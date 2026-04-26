"""
High-level Google Slides audit deck generator.

Exposes a single MCP tool, `create_audit_presentation`, that takes a structured `deck` JSON
payload and produces a fully-built, branded Google Slides presentation:

  1. Copies a template presentation (theme/colors/fonts/master are inherited).
  2. (Optional) Creates a hidden Google Sheet to host data + native charts.
  3. Builds slides programmatically via createSlide / createShape / createTable / createImage
     / createSheetsChart, chunked to stay under Google API per-batch limits.
  4. Adds speaker notes in a second pass.
  5. Places the final deck (and data sheet) into a Drive folder path
     (e.g. ["CLIENTS", "edaa.fr", "SEO"]).
  6. On failure after the copy step, best-effort deletes the partial deck + sheet before
     re-raising, so retries don't accumulate orphans.

See README.md (Slides section) for the full `deck` JSON schema and an example.
"""
from __future__ import annotations

import asyncio
import json
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from googleapiclient.errors import HttpError

from auth.service_decorator import require_multiple_services
from core.server import server
from core.utils import handle_http_errors

from gslides import _builders as B

logger = logging.getLogger(__name__)

# Slides `batchUpdate` becomes unreliable when many `createSlide` calls are
# stacked in one batch alongside dependent `insertText` / `createTable` /
# `replaceImage` requests — Google's backend regularly returns deterministic
# HTTP 500s on such mixed batches for decks of ~25+ slides. We therefore split
# the build into two homogeneous phases (creation, then content), GROUP CONTENT
# REQUESTS PER SLIDE (so a single batch never spans multiple slides), and cap
# batch size aggressively. 10 requests per batch is conservative but rock-solid
# even for 100-slide decks against custom-layout templates.
MAX_REQUESTS_PER_BATCH = 10
# Per-call addChart batches stay comfortably under Sheets quota.
MAX_CHARTS_PER_SHEETS_BATCH = 20
# Inter-batch pause to avoid bursting Slides/Sheets quota over many chunks.
# Google enforces "Write requests per minute per user" = 60 (default) on
# slides.googleapis.com. With Phase A (1 batchUpdate per slide) + Phase B
# (≥1 batchUpdate per slide) a 33-slide deck issues 66+ write requests; if we
# fire them back-to-back the second half hits HTTP 429. Pacing at ~1s between
# batches keeps a steady cadence comfortably under the per-minute limit on
# typical decks. Larger decks may still hit the cap; raising the per-user
# Slides quota is the proper fix for >60-slide builds.
INTER_BATCH_SLEEP_S = 1.0
# Per-batch retry policy. Total worst-case wait per batch ≈ 1+2+4+8+16+32 = 63s.
# Transient Slides 500s nearly always clear within this window.
BATCH_MAX_ATTEMPTS = 6
BATCH_BASE_DELAY_S = 1.0
BATCH_RETRYABLE_STATUSES = {429, 500, 502, 503, 504}

_BASIC_CHART_TYPES = {"BAR", "COLUMN", "LINE", "AREA", "SCATTER", "COMBO", "STEPPED_AREA"}
_PIE_CHART_TYPES = {"PIE", "DOUGHNUT"}

# Style fields that flow from deck.chart_defaults into each chart, or that a single chart
# can override directly in its `chart` block. Anything not in this list is ignored.
_CHART_STYLE_FIELDS = (
    "series_colors",
    "background_color",
    "font_family",
    "title_text_format",
    "legend_position",
    "stacked_type",
)


def _hex_to_rgb_color(hex_color: str) -> Dict[str, float]:
    """'#1A73E8' -> {'red': 0.10..., 'green': 0.45..., 'blue': 0.91...}."""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise Exception(f"Invalid HEX color '{hex_color}'. Expected #RRGGBB.")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}


def _build_text_format(spec: Dict[str, Any]) -> Dict[str, Any]:
    """Translate a friendly text-format dict into the Sheets API's TextFormat shape.

    Accepted keys: bold, italic, font_size, font_family, foreground_color (HEX).
    """
    out: Dict[str, Any] = {}
    if spec.get("bold") is not None:
        out["bold"] = bool(spec["bold"])
    if spec.get("italic") is not None:
        out["italic"] = bool(spec["italic"])
    if spec.get("font_size") is not None:
        out["fontSize"] = int(spec["font_size"])
    if spec.get("font_family"):
        out["fontFamily"] = str(spec["font_family"])
    if spec.get("foreground_color"):
        out["foregroundColorStyle"] = {"rgbColor": _hex_to_rgb_color(spec["foreground_color"])}
    return out


def _effective_chart_style(
    chart_defaults: Optional[Dict[str, Any]], chart_spec: Dict[str, Any]
) -> Dict[str, Any]:
    """Merge per-chart style on top of deck-level chart_defaults. Per-chart wins per-key."""
    style: Dict[str, Any] = {}
    for key in _CHART_STYLE_FIELDS:
        if chart_defaults and key in chart_defaults and chart_defaults[key] is not None:
            style[key] = chart_defaults[key]
        if key in chart_spec and chart_spec[key] is not None:
            style[key] = chart_spec[key]
    return style


async def _run_with_transient_retry(
    call,
    label: str,
    max_attempts: int = BATCH_MAX_ATTEMPTS,
    base_delay: float = BATCH_BASE_DELAY_S,
):
    """Execute `call` (a zero-arg synchronous callable) on a worker thread, retrying
    transient Google API errors locally with exponential backoff.

    This is the per-batch retry layer used inside the audit tool. It absorbs
    transient 5xx and 429 responses on a single batch so the OUTER tool-level
    retry never has to restart the whole presentation build (which would re-copy
    the template, re-create the data sheet, and re-render every slide).

    Non-transient errors (400, 401, 403, 404, etc.) are re-raised immediately.
    After exhausting `max_attempts` on transient errors, the last HttpError is
    re-raised so the caller can surface a clean message.
    """
    last_error: Optional[HttpError] = None
    for attempt in range(max_attempts):
        try:
            return await asyncio.to_thread(call)
        except HttpError as err:
            status = getattr(err.resp, "status", None)
            if status not in BATCH_RETRYABLE_STATUSES:
                raise
            last_error = err
            if attempt >= max_attempts - 1:
                break
            # 429 is rate-limit (per-minute quota), not flaky 5xx. Exponential
            # 1-2-4-8s backoff doesn't actually clear the per-minute window —
            # we need to wait long enough for the rolling minute to refresh.
            # Use a much larger base for 429.
            if status == 429:
                # 5s, 10s, 20s, 40s, 80s, 160s — first retry alone covers most
                # bursts; later retries cover sustained pressure.
                effective_base = max(base_delay, 5.0)
            else:
                effective_base = base_delay
            delay = effective_base * (2 ** attempt)
            logger.warning(
                f"[create_audit_presentation:{label}] Transient HTTP {status} on "
                f"attempt {attempt + 1}/{max_attempts}. Retrying in {delay:.1f}s..."
            )
            await asyncio.sleep(delay)
    assert last_error is not None
    raise last_error


# -----------------------------
# Drive / template helpers
# -----------------------------
async def _copy_template(
    drive_service, template_presentation_id: str, new_title: str
) -> Dict[str, str]:
    """Copy a template Slides file. Returns the new file's metadata."""
    body = {"name": new_title}
    return await _run_with_transient_retry(
        drive_service.files()
        .copy(
            fileId=template_presentation_id,
            body=body,
            fields="id, name, parents, webViewLink",
            supportsAllDrives=True,
        )
        .execute,
        label="drive.files.copy (template)",
    )


async def _delete_file_safe(drive_service, file_id: Optional[str], context: str) -> None:
    if not file_id:
        return
    try:
        await asyncio.to_thread(
            drive_service.files()
            .delete(fileId=file_id, supportsAllDrives=True)
            .execute
        )
        logger.info(f"[create_audit_presentation] Rolled back {context}: deleted {file_id}")
    except Exception as cleanup_err:
        logger.warning(
            f"[create_audit_presentation] Best-effort cleanup of {context} ({file_id}) failed: "
            f"{cleanup_err}"
        )


async def _find_existing_deck_in_folder(
    drive_service, folder_id: str, title: str
) -> Optional[Dict[str, Any]]:
    """Return the most recent presentation with the given exact title in the folder, if any."""
    escaped_title = title.replace("'", "\\'")
    query = (
        f"mimeType='application/vnd.google-apps.presentation' "
        f"and name='{escaped_title}' and '{folder_id}' in parents and trashed=false"
    )
    result = await asyncio.to_thread(
        drive_service.files()
        .list(
            q=query,
            pageSize=5,
            fields="files(id, name, webViewLink, modifiedTime)",
            orderBy="modifiedTime desc",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute
    )
    files = result.get("files", []) or []
    return files[0] if files else None


# -----------------------------
# Sheets data + chart helpers
# -----------------------------
def _column_letter(idx: int) -> str:
    """0-indexed column index -> A1 letter (0 -> A, 25 -> Z, 26 -> AA)."""
    s = ""
    n = idx
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _build_chart_addchart_request(
    sheet_id: int,
    chart_spec: Dict[str, Any],
    data_rows: int,
    data_columns: int,
    style: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Build a single `addChart` request given a chart spec, geometry, and optional style.

    `style` is the merged result of deck.chart_defaults + per-chart overrides, see
    `_effective_chart_style`.
    """
    style = style or {}
    chart_type = chart_spec.get("type", "BAR").upper()
    title = chart_spec.get("title")

    domain_source = {
        "sheetId": sheet_id,
        "startRowIndex": 0,
        "endRowIndex": data_rows,
        "startColumnIndex": 0,
        "endColumnIndex": 1,
    }
    series_sources = [
        {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": data_rows,
            "startColumnIndex": col,
            "endColumnIndex": col + 1,
        }
        for col in range(1, data_columns)
    ]

    spec: Dict[str, Any] = {}
    if title:
        spec["title"] = title

    legend_position = (
        chart_spec.get("legend_position") or style.get("legend_position") or "BOTTOM_LEGEND"
    )

    if chart_type in _BASIC_CHART_TYPES:
        # Horizontal BAR charts have their value axis at the bottom and the
        # category (domain) axis at the left — the opposite of every other
        # basic chart. Series must target BOTTOM_AXIS or Sheets rejects the
        # request with "Bar charts series may only target the BOTTOM_AXIS".
        is_horizontal_bar = chart_type == "BAR"
        domain_axis_position = "LEFT_AXIS" if is_horizontal_bar else "BOTTOM_AXIS"
        value_axis_position = "BOTTOM_AXIS" if is_horizontal_bar else "LEFT_AXIS"

        axis: List[Dict[str, Any]] = []
        if chart_spec.get("domain_axis_title"):
            axis.append({"position": domain_axis_position, "title": chart_spec["domain_axis_title"]})
        if chart_spec.get("value_axis_title"):
            axis.append({"position": value_axis_position, "title": chart_spec["value_axis_title"]})

        series_palette = style.get("series_colors") or []
        # COMBO charts require each series to specify its own type. Without
        # that, Sheets rejects the request with "No basic chart type specified"
        # (a misleading wording that really refers to per-series type).
        # Default mapping: first series = COLUMN, others = LINE. Callers may
        # override via `chart.series_types: ["COLUMN", "LINE", ...]`.
        per_series_types: List[str] = []
        if chart_type == "COMBO":
            override = chart_spec.get("series_types") or []
            for idx in range(len(series_sources)):
                if idx < len(override) and override[idx]:
                    per_series_types.append(str(override[idx]).upper())
                else:
                    per_series_types.append("COLUMN" if idx == 0 else "LINE")

        series_list: List[Dict[str, Any]] = []
        for idx, src in enumerate(series_sources):
            series_entry: Dict[str, Any] = {
                "series": {"sourceRange": {"sources": [src]}},
                "targetAxis": value_axis_position,
            }
            if per_series_types:
                series_entry["type"] = per_series_types[idx]
            if series_palette:
                color_hex = series_palette[idx % len(series_palette)]
                series_entry["colorStyle"] = {"rgbColor": _hex_to_rgb_color(color_hex)}
            series_list.append(series_entry)

        basic_chart: Dict[str, Any] = {
            "chartType": chart_type,
            "legendPosition": legend_position,
            "headerCount": 1,
            "domains": [{"domain": {"sourceRange": {"sources": [domain_source]}}}],
            "series": series_list,
        }
        if axis:
            basic_chart["axis"] = axis
        if style.get("stacked_type"):
            basic_chart["stackedType"] = str(style["stacked_type"]).upper()
        spec["basicChart"] = basic_chart
    elif chart_type in _PIE_CHART_TYPES:
        spec["pieChart"] = {
            "legendPosition": legend_position,
            "domain": {"sourceRange": {"sources": [domain_source]}},
            "series": {"sourceRange": {"sources": [series_sources[0]]}},
        }
        if chart_type == "DOUGHNUT":
            spec["pieChart"]["pieHole"] = 0.5
    else:
        raise Exception(
            f"Unsupported chart type '{chart_type}'. "
            f"Supported: {sorted(_BASIC_CHART_TYPES | _PIE_CHART_TYPES)}"
        )

    if style.get("background_color"):
        spec["backgroundColorStyle"] = {"rgbColor": _hex_to_rgb_color(style["background_color"])}
    if style.get("font_family"):
        spec["fontName"] = str(style["font_family"])
    if style.get("title_text_format"):
        spec["titleTextFormat"] = _build_text_format(style["title_text_format"])

    return {
        "addChart": {
            "chart": {
                "spec": spec,
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": sheet_id,
                            "rowIndex": 0,
                            "columnIndex": data_columns + 1,
                        },
                        "widthPixels": int(chart_spec.get("width_pixels", 600)),
                        "heightPixels": int(chart_spec.get("height_pixels", 371)),
                    }
                },
            }
        }
    }


async def _create_data_sheet_and_charts(
    sheets_service,
    drive_service,
    title: str,
    target_folder_id: Optional[str],
    chart_specs: List[Dict[str, Any]],
    chart_defaults: Optional[Dict[str, Any]] = None,
) -> Tuple[Dict[str, Any], List[Tuple[str, int]]]:
    """Create one Sheet with one tab per chart, write data, create charts, return (sheet_meta, chart_ids).

    `chart_ids` is a list of (chart_uid, sheets_chart_id) where chart_uid matches the per-slide reference.
    """
    sheet_titles = [f"Chart_{i + 1}" for i in range(len(chart_specs))]

    body = {
        "properties": {"title": f"{title} - data"},
        "sheets": [{"properties": {"title": t}} for t in sheet_titles],
    }
    spreadsheet = await _run_with_transient_retry(
        sheets_service.spreadsheets().create(body=body).execute,
        label="sheets.spreadsheets.create",
    )
    spreadsheet_id = spreadsheet["spreadsheetId"]
    spreadsheet_url = spreadsheet.get("spreadsheetUrl")
    sheet_id_by_index = {
        i: spreadsheet["sheets"][i]["properties"]["sheetId"]
        for i in range(len(spreadsheet["sheets"]))
    }

    # Move the data sheet next to the deck if a target folder was resolved.
    if target_folder_id:
        from gdrive.drive_helpers import move_file_to_folder

        await move_file_to_folder(
            drive_service, spreadsheet_id, target_folder_id, file_name=f"{title} - data"
        )

    # Phase 1: write data with one values.batchUpdate.
    value_ranges = []
    geometry: List[Tuple[int, int, int]] = []  # (chart_index, data_rows, data_columns)
    for i, spec in enumerate(chart_specs):
        data = spec.get("data") or {}
        headers = data.get("headers") or []
        rows = data.get("rows") or []
        if not headers or not rows:
            raise Exception(
                f"Chart spec #{i + 1} ('{spec.get('title') or 'untitled'}') is missing "
                "data.headers or data.rows."
            )

        # Build a 2D grid of strings/numbers.
        grid = [list(headers)] + [list(r) for r in rows]
        n_rows = len(grid)
        n_cols = len(headers)
        end_col_letter = _column_letter(n_cols - 1)
        sheet_title = sheet_titles[i]
        a1_range = f"'{sheet_title}'!A1:{end_col_letter}{n_rows}"
        value_ranges.append({"range": a1_range, "values": grid})
        geometry.append((i, n_rows, n_cols))

    await _run_with_transient_retry(
        sheets_service.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": "USER_ENTERED", "data": value_ranges},
        )
        .execute,
        label="sheets.values.batchUpdate",
    )

    # Phase 2: addChart, chunked to stay under Sheets per-batch limits.
    addchart_requests = [
        _build_chart_addchart_request(
            sheet_id=sheet_id_by_index[chart_idx],
            chart_spec=chart_specs[chart_idx],
            data_rows=n_rows,
            data_columns=n_cols,
            style=_effective_chart_style(chart_defaults, chart_specs[chart_idx]),
        )
        for (chart_idx, n_rows, n_cols) in geometry
    ]

    chart_ids: List[Tuple[str, int]] = []
    chunks = B.chunk_requests(addchart_requests, MAX_CHARTS_PER_SHEETS_BATCH)
    chunk_offset = 0
    for chunk_idx, chunk in enumerate(chunks):
        response = await _run_with_transient_retry(
            sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": chunk})
            .execute,
            label=f"sheets.addChart batch {chunk_idx + 1}/{len(chunks)} ({len(chunk)} charts)",
        )
        replies = response.get("replies", [])
        for j, reply in enumerate(replies):
            if "addChart" not in reply:
                raise Exception("Sheets API returned an unexpected reply (missing addChart).")
            chart_uid = chart_specs[chunk_offset + j].get("_uid", f"chart_{chunk_offset + j}")
            chart_ids.append((chart_uid, reply["addChart"]["chart"]["chartId"]))
        chunk_offset += len(chunk)
        if chunk_idx < len(chunks) - 1:
            await asyncio.sleep(INTER_BATCH_SLEEP_S)

    sheet_meta = {
        "spreadsheet_id": spreadsheet_id,
        "spreadsheet_url": spreadsheet_url,
        "chart_count": len(chart_ids),
    }
    return sheet_meta, chart_ids


# -----------------------------
# Slides building
# -----------------------------
async def _strip_existing_slides(slides_service, presentation_id: str) -> None:
    """Delete every slide currently in the (just-copied) presentation."""
    presentation = await _run_with_transient_retry(
        slides_service.presentations().get(presentationId=presentation_id).execute,
        label="slides.presentations.get (strip)",
    )
    slides = presentation.get("slides", []) or []
    if not slides:
        return
    requests = [{"deleteObject": {"objectId": s["objectId"]}} for s in slides]
    await _run_with_transient_retry(
        slides_service.presentations()
        .batchUpdate(presentationId=presentation_id, body={"requests": requests})
        .execute,
        label="slides.batchUpdate (strip existing slides)",
    )


async def _execute_slides_batches(
    slides_service, presentation_id: str, requests: List[Dict[str, Any]]
) -> None:
    """Run a flat list of Slides requests in chunks of MAX_REQUESTS_PER_BATCH.

    Each chunk has its own transient-error retry layer with exponential backoff,
    so a single 500/503 on chunk N does not abort the whole tool — only that
    chunk is retried. Chunks are paced by INTER_BATCH_SLEEP_S to avoid 429s on
    very large decks.
    """
    if not requests:
        return
    chunks = B.chunk_requests(requests, MAX_REQUESTS_PER_BATCH)
    total = len(chunks)
    for chunk_idx, chunk in enumerate(chunks):
        await _run_with_transient_retry(
            slides_service.presentations()
            .batchUpdate(presentationId=presentation_id, body={"requests": chunk})
            .execute,
            label=f"slides.batchUpdate {chunk_idx + 1}/{total} ({len(chunk)} requests)",
        )
        if chunk_idx < total - 1:
            await asyncio.sleep(INTER_BATCH_SLEEP_S)


# Slides request types whose payload references an OBJECT THAT MUST ALREADY
# EXIST on the slide (typically a placeholder objectId we pre-allocated via
# `placeholderIdMappings`). If that object never materialized — which happens
# silently when a `layoutPlaceholder` mapping doesn't match the layout's real
# placeholders — the request will fail. Slides returns generic HTTP 500 for
# this case (instead of a proper 400), so we filter these requests out
# defensively after Phase A using a refetched view of the presentation.
_OBJECT_REFERENCING_REQUEST_FIELDS: Dict[str, str] = {
    "insertText": "objectId",
    "deleteText": "objectId",
    "updateTextStyle": "objectId",
    "updateParagraphStyle": "objectId",
    "createParagraphBullets": "objectId",
    "deleteParagraphBullets": "objectId",
    "updateShapeProperties": "objectId",
    "updateImageProperties": "objectId",
    "replaceImage": "imageObjectId",
    "deleteObject": "objectId",
}

# Phase B request types that CREATE a new pageElement with a custom objectId
# we provide. These objects don't exist yet at verification time (Phase B
# hasn't run) but the dependent requests (e.g. cell insertText into a freshly
# created table) DO target them — so we must whitelist these IDs as
# "will-exist" when filtering.
_OBJECT_CREATING_REQUEST_KINDS: Tuple[str, ...] = (
    "createTable",
    "createShape",
    "createImage",
    "createLine",
    "createVideo",
    "createSheetsChart",
)


def _extract_target_object_id(request: Dict[str, Any]) -> Optional[str]:
    """Return the objectId a Slides request expects to ALREADY exist, or None.

    Only used to filter Phase B requests against a live view of the
    presentation. Requests that *create* a new object (createShape,
    createTable, createImage, createSheetsChart, createSlide) reference
    the slide via `pageObjectId`, which is always pre-allocated by us and
    confirmed to exist after Phase A — those are intentionally not filtered.
    """
    for kind, field in _OBJECT_REFERENCING_REQUEST_FIELDS.items():
        if kind in request:
            return (request[kind] or {}).get(field)
    return None


def _rebind_request_object_id(
    request: Dict[str, Any], rebind_map: Dict[str, str]
) -> Dict[str, Any]:
    """Rewrite the `objectId` (or `imageObjectId`) field of a Slides request
    using `rebind_map` if its current value is a pseudo objectId we deferred.

    Returns a NEW request dict if rewriting happened, or the original request
    untouched. Does not mutate input.

    Only the top-level `objectId` of object-referencing requests is rewritten:
    `insertText`, `updateTextStyle`, `updateParagraphStyle`,
    `updateShapeProperties`, `updatePageElementTransform`, `replaceImage`,
    `deleteObject`. Create-new requests (createTable, createShape, etc.)
    keep their newly-minted IDs untouched.
    """
    for kind, field in _OBJECT_REFERENCING_REQUEST_FIELDS.items():
        if kind not in request:
            continue
        inner = request[kind] or {}
        old_id = inner.get(field)
        if old_id and old_id in rebind_map:
            new_inner = dict(inner)
            new_inner[field] = rebind_map[old_id]
            return {kind: new_inner}
        return request
    return request


def _extract_created_object_id(request: Dict[str, Any]) -> Optional[str]:
    """Return the custom objectId that a Phase B create* request will mint.

    These IDs don't appear in the live presentation yet (the create has not
    been issued), but downstream requests in the same Phase B run target
    them. We must NOT filter those downstream requests — that would orphan
    cell `insertText` calls that depend on a `createTable` queued later in
    the same Phase B batch sequence.
    """
    for kind in _OBJECT_CREATING_REQUEST_KINDS:
        if kind in request:
            return (request[kind] or {}).get("objectId")
    return None


def _collect_existing_object_ids(presentation: Dict[str, Any]) -> set:
    """Return every pageElement objectId currently materialized in the deck."""
    out: set = set()
    for page in presentation.get("slides", []) or []:
        for el in page.get("pageElements", []) or []:
            oid = el.get("objectId")
            if oid:
                out.add(oid)
    return out


async def _execute_slides_per_slide(
    slides_service,
    presentation_id: str,
    per_slide_requests: List[Tuple[int, List[Dict[str, Any]]]],
) -> None:
    """Run Slides content requests one slide at a time.

    Each slide's content requests are sent in their own ``batchUpdate`` call
    (chunked to ``MAX_REQUESTS_PER_BATCH`` if the slide has many cells / heavy
    content like big tables). Crucially, a single ``batchUpdate`` is never
    allowed to mix requests targeting different slides — that combination is
    what reliably triggers ``HTTP 500: Internal error encountered`` against
    custom-layout templates, even for content-only requests.

    Each chunk goes through the per-call transient-retry helper, so a flake on
    slide N does not abort the whole tool.
    """
    if not per_slide_requests:
        return
    # Fail-fast retry budgets so the per-request fallback ALWAYS runs within
    # the MCP transport's 60s timeout window. Worst-case math:
    #   chunk: 2 attempts × (≈5s API + 1s sleep) ≈ 12s
    #   fallback: N requests × 3 attempts × (≈5s API + ≤4s sleep) per request
    # For typical slides with ≤10 content requests, total well under 60s.
    # We deliberately do NOT use the generic 6-attempt budget (1+2+4+8+16+32
    # = 63s of sleeps alone) which prevented the fallback from ever firing
    # in practice.
    chunk_max_attempts = 2
    per_request_max_attempts = 3
    total_slides = len(per_slide_requests)
    for slide_pos, (slide_idx, slide_requests) in enumerate(per_slide_requests):
        if not slide_requests:
            continue
        chunks = B.chunk_requests(slide_requests, MAX_REQUESTS_PER_BATCH)
        for chunk_idx, chunk in enumerate(chunks):
            label = (
                f"slides.batchUpdate slide#{slide_idx + 1} "
                f"({slide_pos + 1}/{total_slides}) chunk {chunk_idx + 1}/{len(chunks)} "
                f"({len(chunk)} requests)"
            )
            try:
                await _run_with_transient_retry(
                    slides_service.presentations()
                    .batchUpdate(presentationId=presentation_id, body={"requests": chunk})
                    .execute,
                    label=label,
                    max_attempts=chunk_max_attempts,
                )
            except HttpError as e:
                # The Slides backend has a known pathology: certain content
                # batches against custom layouts (notably multi-BODY two-column
                # layouts) deterministically return HTTP 500 even when every
                # individual request is valid. Sending each request in its
                # own batchUpdate sidesteps that — and as a bonus, if a
                # request is genuinely broken, the per-request error pinpoints
                # which one. We only fall back on 5xx, not 4xx (the latter
                # are real client errors that won't get better by splitting).
                status = getattr(e, "status_code", None) or (
                    getattr(e, "resp", None).status if getattr(e, "resp", None) else None
                )
                if status not in (500, 502, 503, 504):
                    raise
                logger.warning(
                    f"[create_audit_presentation:{label}] Batch failed with HTTP {status} "
                    f"after retries; falling back to per-request execution to isolate the "
                    f"problem and work around the Slides backend's multi-request 500 pathology."
                )
                for req_idx, single_req in enumerate(chunk):
                    sub_label = (
                        f"slides.batchUpdate slide#{slide_idx + 1} "
                        f"({slide_pos + 1}/{total_slides}) chunk {chunk_idx + 1}/{len(chunks)} "
                        f"req {req_idx + 1}/{len(chunk)} [{next(iter(single_req.keys()))}]"
                    )
                    try:
                        await _run_with_transient_retry(
                            slides_service.presentations()
                            .batchUpdate(
                                presentationId=presentation_id,
                                body={"requests": [single_req]},
                            )
                            .execute,
                            label=sub_label,
                            max_attempts=per_request_max_attempts,
                        )
                    except HttpError as sub_e:
                        sub_status = getattr(sub_e, "status_code", None) or (
                            getattr(sub_e, "resp", None).status
                            if getattr(sub_e, "resp", None)
                            else None
                        )
                        # Truthful diagnostic: dump the request that we're
                        # giving up on so the user can fix or remove it.
                        try:
                            req_dump = json.dumps(single_req, ensure_ascii=False)[:1500]
                        except Exception:
                            req_dump = repr(single_req)[:1500]
                        if sub_status in (500, 502, 503, 504):
                            logger.error(
                                f"[create_audit_presentation:{sub_label}] Slides API still "
                                f"returns HTTP {sub_status} for this single request after "
                                f"the per-batch retry budget. Skipping it so the rest of "
                                f"the deck can complete. Request body: {req_dump}"
                            )
                            continue
                        # 4xx and other errors are real bugs in the request
                        # — surface them so they get fixed.
                        logger.error(
                            f"[create_audit_presentation:{sub_label}] Slides API rejected "
                            f"this request with HTTP {sub_status}. Request body: {req_dump}"
                        )
                        raise
                    await asyncio.sleep(INTER_BATCH_SLEEP_S)
            await asyncio.sleep(INTER_BATCH_SLEEP_S)


async def _execute_slides_sequentially(
    slides_service,
    presentation_id: str,
    requests: List[Dict[str, Any]],
    label_prefix: str,
) -> None:
    """Run each Slides request as its own `batchUpdate` call.

    This is the bulletproof path for `createSlide` requests that target custom
    or custom-derived layouts (e.g. layouts whose internal `name` is something
    like `TITLE_AND_BODY_1_2`). The Slides backend reliably 500s
    ('Internal error encountered') when several such `createSlide` calls share
    a single `batchUpdate`, even when the requests themselves are individually
    valid and even at a batch size as small as 30. Sending them one by one
    sidesteps that class of issue entirely while keeping the per-call retry
    layer for genuine transient errors.

    Each call still uses `_run_with_transient_retry`, so transient 5xx/429 on
    a single create get absorbed locally without restarting the build.
    """
    if not requests:
        return
    total = len(requests)
    for i, req in enumerate(requests):
        await _run_with_transient_retry(
            slides_service.presentations()
            .batchUpdate(presentationId=presentation_id, body={"requests": [req]})
            .execute,
            label=f"{label_prefix} {i + 1}/{total}",
        )
        if i < total - 1:
            await asyncio.sleep(INTER_BATCH_SLEEP_S)


def _annotate_chart_uids(deck: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Walk the deck, assign a unique _uid to every chart spec, and return the flat chart list."""
    flat: List[Dict[str, Any]] = []
    for i, slide in enumerate(deck.get("slides") or []):
        chart = slide.get("chart")
        if chart:
            uid = f"slide{i}_chart"
            chart["_uid"] = uid
            flat.append(chart)
    return flat


# -----------------------------
# Public MCP tool
# -----------------------------
@server.tool()
@handle_http_errors("create_audit_presentation", service_type="slides")
@require_multiple_services(
    [
        {"service_type": "slides", "scopes": "slides", "param_name": "slides_service"},
        {"service_type": "sheets", "scopes": "sheets_write", "param_name": "sheets_service"},
        {"service_type": "drive", "scopes": "drive_full", "param_name": "drive_service"},
    ]
)
async def create_audit_presentation(
    slides_service,
    sheets_service,
    drive_service,
    user_google_email: str,
    template_presentation_id: str,
    deck: Dict[str, Any],
    folder_id: Optional[str] = None,
    folder_path: Optional[List[str]] = None,
    create_folders_if_missing: bool = True,
    if_exists: str = "create_new",
    cleanup_data_sheet: bool = False,
    keep_template_slides: bool = False,
    keep_on_error: bool = False,
) -> str:
    """
    Build a full Google Slides deck from a structured JSON payload, copying a template for branding.

    Workflow (one MCP call -> whole deck):
      1) Copy the template (theme, masters, fonts, colors are inherited).
      2) If any slide needs a chart, create a hidden data Sheet, write the data, create native
         Sheets charts, and capture their chartIds.
      3) Build slides programmatically (createSlide + placeholders + tables + images +
         createSheetsChart). Requests are chunked internally to stay under Google API limits.
      4) Add speaker notes in a second pass (requires reading the presentation back to find each
         slide's speakerNotesObjectId).
      5) Move the final deck (and data sheet) to the target Drive folder.
      6) On failure after the copy step, best-effort delete the partial files before re-raising.

    `deck` JSON shape:
      {
        "title": "Pre-audit SEO - edaa.fr - 2026-04",
        "chart_defaults": {                           # optional, applied to every chart
          "series_colors": ["#1A73E8", "#34A853", "#FBBC04", "#EA4335"],
          "background_color": "#FFFFFF",
          "font_family": "Roboto",
          "title_text_format": {"bold": true, "font_size": 14, "foreground_color": "#202124"},
          "legend_position": "BOTTOM_LEGEND",
          "stacked_type": "STACKED"                   # NONE | STACKED | PERCENT_STACKED
        },
        "slides": [
          {"layout": "TITLE", "fields": {"title": "...", "subtitle": "..."}},
          {"layout": "TITLE_AND_BODY", "fields": {"title": "...", "body": "..."},
           "speaker_notes": "..."},
          # Two-column layout: pass `body` as a list. Item 0 fills BODY index 0
          # (left column), item 1 fills BODY index 1 (right column), etc.
          {"layout": "Title + Two Columns", "fields": {
              "title": "Avant / Après",
              "body": ["Avant: ...", "Après: ..."]}},
          # PICTURE placeholder ("espace réservé image"): fill by ordinal
          # index of the PICTURE placeholders in the layout. Each item is
          # either a URL string or {"url": "...", "method": "CENTER_INSIDE"|"CENTER_CROP"}.
          {"layout": "Cover", "fields": {"title": "Pré-audit SEO"},
           "image_placeholders": ["https://example.com/laptop-mockup.png"]},
          {"layout": "BLANK", "title": "Free title", "table": {
              "headers": ["Metric", "Value"], "rows": [["...", "..."]]}},
          {"layout": "BLANK", "title": "Scores", "chart": {
              "type": "COLUMN", "title": "Scores par pilier",
              "data": {"headers": ["Pilier", "Score"],
                       "rows": [["Technique", 90], ["Contenu", 64]]}}},
          {"layout": "SECTION_HEADER", "fields": {"title": "Recommandations"}},
          {"layout": "BLANK", "title": "Capture", "image": {"url": "https://..."}}
        ]
      }

    Supported slide layouts (predefined): BLANK, TITLE, TITLE_AND_BODY, TITLE_AND_TWO_COLUMNS,
    TITLE_ONLY, SECTION_HEADER, SECTION_TITLE_AND_DESCRIPTION, ONE_COLUMN_TEXT, MAIN_POINT,
    BIG_NUMBER. Custom layout display names from the template are also resolved.

    Supported chart types: BAR, COLUMN, LINE, AREA, SCATTER, COMBO, STEPPED_AREA, PIE, DOUGHNUT.

    Args:
        user_google_email: The user's Google email address. Required.
        template_presentation_id: Drive file ID of a Google Slides template to copy. Required.
        deck: Structured deck definition (see schema above). Required.
        folder_id: Specific Drive folder ID to place the deck in. Mutually exclusive with folder_path.
        folder_path: Folder path to navigate/create (e.g. ["CLIENTS", "edaa.fr", "SEO"]).
        create_folders_if_missing: If True, missing folders in folder_path are created. Default True.
        if_exists: One of "create_new" (append timestamp suffix), "replace" (delete existing in
            folder), or "skip" (return existing deck untouched). Default "create_new".
        cleanup_data_sheet: If True, delete the data Sheet after the deck is built. Default False
            so charts remain refreshable.
        keep_template_slides: If True, the slides already present in the copied template are
            preserved and the generated slides are appended AFTER them. Default False (the
            template copy is wiped first so the deck only contains generated slides). Use this
            when the template ships with fixed boilerplate (cover, methodology, about-us, ...)
            that should appear in every audit deck.
        keep_on_error: If True, the partial deck (and data sheet) are NOT deleted when the build
            fails. Use this for debugging — open the half-built file in Drive to see exactly
            which slides made it and inspect the resulting placeholder/objectId state. Default
            False so production retries don't accumulate orphan files. The presentation_id of
            the kept partial is logged at WARNING level so it's easy to find.

    Returns:
        str: JSON string with presentation_id, presentation_url, slide_count, data_sheet_url,
        folder_id, message.
    """
    if not isinstance(template_presentation_id, str) or not template_presentation_id.strip():
        raise Exception(
            "'template_presentation_id' is required and must be a non-empty Drive file ID "
            "(the long token in the template's URL between '/d/' and '/edit')."
        )
    template_presentation_id = template_presentation_id.strip()
    if not isinstance(deck, dict):
        raise Exception("'deck' must be a JSON object with 'title' and 'slides'.")
    deck_title = (deck.get("title") or "").strip()
    if not deck_title:
        raise Exception("'deck.title' is required.")
    slides = deck.get("slides") or []
    if not slides:
        raise Exception("'deck.slides' must be a non-empty list.")
    if if_exists not in {"create_new", "replace", "skip"}:
        raise Exception("'if_exists' must be one of: create_new, replace, skip.")

    logger.info(
        f"[create_audit_presentation] Email: '{user_google_email}', Template: "
        f"{template_presentation_id}, Deck title: '{deck_title}', Slides: {len(slides)}"
    )

    # 0) Pre-flight: read the TEMPLATE's layouts directly and validate every
    # slide's `layout` against them before copying anything. This way a typo
    # like 'Title + Chart + Body' (when the template only has 'Title + Chart')
    # fails immediately with a clear error listing the available layouts —
    # without leaving an orphan copy in Drive.
    template_meta = await _run_with_transient_retry(
        slides_service.presentations().get(presentationId=template_presentation_id).execute,
        label="slides.presentations.get (template preflight)",
    )

    # Trace EVERY custom layout the Slides API exposes for this template,
    # along with which master each one belongs to. This is the single most
    # useful diagnostic when Google complains 'predefined layout (X) is not
    # present in the current master': the log will show exactly what was
    # actually copied vs. what the editor UI shows. Use repr() to surface any
    # invisible Unicode (NBSPs, zero-width spaces, etc.) drifting into a
    # layout name.
    template_masters = [
        {"objectId": m.get("objectId"), "displayName": (m.get("masterProperties") or {}).get("displayName")}
        for m in (template_meta.get("masters") or [])
    ]
    template_layouts_trace = []
    for layout in template_meta.get("layouts") or []:
        props = layout.get("layoutProperties") or {}
        template_layouts_trace.append(
            {
                "objectId": layout.get("objectId"),
                "displayName": props.get("displayName"),
                "name": props.get("name"),
                "masterObjectId": props.get("masterObjectId"),
            }
        )
    logger.info(
        f"[create_audit_presentation] Template preflight: "
        f"{len(template_masters)} master(s), {len(template_layouts_trace)} layout(s)."
    )
    logger.info(
        f"[create_audit_presentation] Template masters: "
        f"{json.dumps(template_masters, ensure_ascii=False)}"
    )
    logger.info(
        f"[create_audit_presentation] Template layouts (use repr() output to spot hidden "
        f"Unicode in displayName): "
        f"{json.dumps(template_layouts_trace, ensure_ascii=False)}"
    )
    # Also log the EXACT slide layout names from the deck JSON for side-by-side
    # comparison with the template_layouts_trace dump above.
    requested_layouts = sorted({(s.get("layout") or "BLANK") for s in slides})
    logger.info(
        f"[create_audit_presentation] Deck requests {len(requested_layouts)} distinct layout "
        f"name(s): {json.dumps(requested_layouts, ensure_ascii=False)}"
    )

    layout_errors: List[str] = []
    for i, slide_spec in enumerate(slides):
        try:
            B.resolve_layout_reference(template_meta, slide_spec.get("layout") or "BLANK")
        except Exception as resolve_err:
            layout_errors.append(f"  • Slide #{i + 1}: {resolve_err}")
    if layout_errors:
        raise Exception(
            "Layout validation failed before any Drive operation:\n"
            + "\n".join(layout_errors)
        )

    # 1) Resolve target folder (if any) BEFORE the copy so we can do `if_exists` checks.
    target_folder_id = folder_id
    folder_path_summary = ""
    if folder_path and not folder_id:
        from gdrive.drive_helpers import find_or_create_folder_path

        folder_result = await find_or_create_folder_path(
            drive_service,
            folder_path,
            root_folder_id=None,
            create_missing=create_folders_if_missing,
        )
        if folder_result:
            target_folder_id = folder_result["id"]
            folder_path_summary = folder_result["path_summary"]
        else:
            logger.warning(
                f"[create_audit_presentation] Could not navigate folder_path "
                f"{' > '.join(folder_path)}; deck will land in My Drive."
            )

    # 2) Handle if_exists policy (only meaningful when we know the target folder).
    final_title = deck_title
    if target_folder_id and if_exists in {"replace", "skip"}:
        existing = await _find_existing_deck_in_folder(
            drive_service, target_folder_id, deck_title
        )
        if existing:
            if if_exists == "skip":
                msg = (
                    f"Skipped: a deck named '{deck_title}' already exists in the target folder."
                )
                return json.dumps(
                    {
                        "presentation_id": existing["id"],
                        "presentation_url": existing.get("webViewLink"),
                        "slide_count": None,
                        "data_sheet_url": None,
                        "folder_id": target_folder_id,
                        "message": msg,
                    },
                    indent=2,
                )
            await _delete_file_safe(drive_service, existing["id"], "existing deck (if_exists=replace)")
    elif if_exists == "create_new":
        # Always disambiguate when there's a folder but a duplicate may exist.
        if target_folder_id:
            existing = await _find_existing_deck_in_folder(
                drive_service, target_folder_id, deck_title
            )
            if existing:
                final_title = f"{deck_title} - {datetime.utcnow().strftime('%Y%m%d-%H%M%S')}"
                logger.info(
                    f"[create_audit_presentation] Duplicate detected; using disambiguated title: "
                    f"'{final_title}'"
                )

    # 3) Copy template.
    new_file = await _copy_template(drive_service, template_presentation_id, final_title)
    presentation_id = new_file["id"]
    presentation_url = (
        new_file.get("webViewLink")
        or f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    )

    data_sheet_meta: Optional[Dict[str, Any]] = None
    try:
        # 4) Move the deck into the target folder (if any).
        if target_folder_id:
            from gdrive.drive_helpers import move_file_to_folder

            await move_file_to_folder(
                drive_service, presentation_id, target_folder_id, file_name=final_title
            )

        # 5) Either wipe the template's slides (clean canvas) or keep them (boilerplate mode).
        if not keep_template_slides:
            await _strip_existing_slides(slides_service, presentation_id)

        # 6) If any slide needs a chart, build the data Sheet first.
        flat_chart_specs = _annotate_chart_uids(deck)
        chart_defaults = deck.get("chart_defaults") or None
        chart_id_by_uid: Dict[str, int] = {}
        if flat_chart_specs:
            data_sheet_meta, chart_pairs = await _create_data_sheet_and_charts(
                sheets_service,
                drive_service,
                title=final_title,
                target_folder_id=target_folder_id,
                chart_specs=flat_chart_specs,
                chart_defaults=chart_defaults,
            )
            chart_id_by_uid = dict(chart_pairs)

        # 7) Build all slide-creation requests in one big list (we'll chunk for execution).
        presentation = await _run_with_transient_retry(
            slides_service.presentations().get(presentationId=presentation_id).execute,
            label="slides.presentations.get (layout discovery)",
        )

        # Trace the COPIED presentation's masters and layouts. If this differs
        # from the template preflight log above, we know `drive.files.copy`
        # didn't preserve everything (rare but possible with PPTX-imported
        # templates) and the user can act on it directly.
        copy_masters = [
            {"objectId": m.get("objectId"),
             "displayName": (m.get("masterProperties") or {}).get("displayName")}
            for m in (presentation.get("masters") or [])
        ]
        copy_layouts_trace = []
        for layout in presentation.get("layouts") or []:
            props = layout.get("layoutProperties") or {}
            copy_layouts_trace.append(
                {
                    "objectId": layout.get("objectId"),
                    "displayName": props.get("displayName"),
                    "name": props.get("name"),
                    "masterObjectId": props.get("masterObjectId"),
                }
            )
        logger.info(
            f"[create_audit_presentation] Copy state after strip: "
            f"{len(copy_masters)} master(s), {len(copy_layouts_trace)} layout(s)."
        )
        logger.info(
            f"[create_audit_presentation] Copy masters: "
            f"{json.dumps(copy_masters, ensure_ascii=False)}"
        )
        logger.info(
            f"[create_audit_presentation] Copy layouts: "
            f"{json.dumps(copy_layouts_trace, ensure_ascii=False)}"
        )
        # Dump the EXACT placeholders each layout exposes (type + index +
        # objectId). This is critical for debugging two-column / multi-body
        # layouts: if our `placeholderIdMappings` references a (type, index)
        # combo the layout doesn't actually expose, the Slides backend
        # silently drops the mapping at createSlide time AND THEN returns
        # generic HTTP 500 on any insertText / replaceImage that targets the
        # never-bound objectId. Logging the truth here makes the cause
        # immediately visible from the server logs.
        layout_placeholders_trace: List[Dict[str, Any]] = []
        for layout in presentation.get("layouts") or []:
            props = layout.get("layoutProperties") or {}
            phs: List[Dict[str, Any]] = []
            for el in layout.get("pageElements") or []:
                placeholder = (el.get("shape") or {}).get("placeholder") or {}
                if not placeholder:
                    continue
                phs.append(
                    {
                        "type": placeholder.get("type"),
                        "index": placeholder.get("index", 0),
                        "objectId": el.get("objectId"),
                    }
                )
            layout_placeholders_trace.append(
                {
                    "displayName": props.get("displayName"),
                    "objectId": layout.get("objectId"),
                    "placeholders": phs,
                }
            )
        logger.info(
            f"[create_audit_presentation] Copy layout placeholders (type/index/objectId): "
            f"{json.dumps(layout_placeholders_trace, ensure_ascii=False)}"
        )

        # Re-resolve every slide's layout against the COPY (not the template)
        # and log the resolution path. If anything resolves to predefinedLayout
        # while we expected a custom layout, we'll see it immediately.
        for index, slide_spec in enumerate(slides):
            requested = slide_spec.get("layout") or "BLANK"
            try:
                ref = B.resolve_layout_reference(presentation, requested)
                if "layoutId" in ref:
                    matched = next(
                        (
                            f"{l['displayName']!r} ({l['objectId']})"
                            for l in copy_layouts_trace
                            if l["objectId"] == ref["layoutId"]
                        ),
                        ref["layoutId"],
                    )
                    logger.info(
                        f"[create_audit_presentation] Slide #{index + 1} layout='{requested}' "
                        f"-> custom layoutId={matched}"
                    )
                else:
                    logger.info(
                        f"[create_audit_presentation] Slide #{index + 1} layout='{requested}' "
                        f"-> predefinedLayout={ref.get('predefinedLayout')}"
                    )
            except Exception as resolve_err:
                # Preflight already validated against the template, so this would
                # only fire if the COPY lost something. Log loudly; the build
                # below will then re-raise via the same resolver call.
                logger.error(
                    f"[create_audit_presentation] Slide #{index + 1} layout='{requested}' "
                    f"FAILED to resolve against the copied presentation: {resolve_err}"
                )

        # When keeping template slides, append generated slides at the end so the boilerplate
        # (cover, methodology, ...) stays in front. Otherwise the template is already empty.
        slide_offset = len(presentation.get("slides", []) or []) if keep_template_slides else 0

        # Two-phase build: ALL `createSlide` requests first (cheap, predictable,
        # no in-batch dependencies), then ALL content requests in a second pass
        # (insertText, updateTextStyle, createTable, createShape, createImage,
        # createSheetsChart, replaceImage). This eliminates the deterministic
        # HTTP 500s Google returns when createSlide and dependent inserts are
        # interleaved in the same batchUpdate.
        creation_requests: List[Dict[str, Any]] = []
        # Per-slide content requests so each batchUpdate stays scoped to a
        # single slide. Mixing content requests for multiple slides in the
        # same batch reliably triggers HTTP 500s against custom layouts.
        per_slide_content: List[Tuple[int, List[Dict[str, Any]]]] = []
        slide_id_index_pairs: List[Tuple[str, int]] = []
        # Aggregated pseudo-objectId → (slide_id, ph_type, layout_index) map.
        # Resolved post-Phase-A by reading the live deck and replacing pseudo
        # IDs in every content request with the real auto-assigned objectId.
        # This is the workaround for the Slides multi-same-type-placeholder
        # binding bug (see build_slide_with_placeholders docstring).
        all_deferred_lookups: Dict[str, Tuple[str, str, int]] = {}

        for index, slide_spec in enumerate(slides):
            (
                slide_id,
                slide_creation,
                slide_content,
                slide_placeholders,
                slide_deferred_lookups,
            ) = B.build_slide_with_placeholders(
                presentation=presentation,
                slide_spec=slide_spec,
                insertion_index=slide_offset + index,
            )
            slide_id_index_pairs.append((slide_id, index))
            creation_requests.extend(slide_creation)
            all_deferred_lookups.update(slide_deferred_lookups)

            this_slide_content: List[Dict[str, Any]] = list(slide_content)

            skipped = slide_placeholders.get("__skipped__")
            if skipped:
                logger.warning(
                    f"[create_audit_presentation] Slide #{index + 1} layout="
                    f"'{slide_spec.get('layout')}': layout has no placeholder for "
                    f"{skipped}; those fields will not be rendered. Add the missing "
                    f"placeholder(s) to the layout in your template, or remove the "
                    f"corresponding field(s) from the slide JSON."
                )

            chart_spec = slide_spec.get("chart")
            if chart_spec:
                uid = chart_spec.get("_uid")
                chart_id = chart_id_by_uid.get(uid)
                if chart_id is None:
                    raise Exception(
                        f"Internal error: chart for slide #{index + 1} was not created."
                    )
                # createSheetsChart targets a slide that must already exist, so
                # it runs in the content phase too.
                this_slide_content.extend(
                    B.build_sheets_chart_requests(
                        slide_id=slide_id,
                        spreadsheet_id=data_sheet_meta["spreadsheet_id"],
                        chart_id=chart_id,
                        position=chart_spec.get("position"),
                    )
                )

            per_slide_content.append((index, this_slide_content))

        # Phase A: create all slides first, ONE PER CALL. Slides batchUpdate
        # is unreliable when several `createSlide` requests targeting custom
        # or derived layouts ride in the same batch — Google returns generic
        # HTTP 500s ('Internal error encountered') even at small batch sizes.
        # Sequential creation is slower (~250-500ms per slide) but rock-solid
        # and lets each individual create get its own retry budget.
        # Distinctive marker so logs make it obvious which code path is live —
        # if you don't see this line in the server logs, your runtime is on a
        # stale build that still does batched createSlide.
        logger.info(
            f"[create_audit_presentation] Phase A: creating {len(creation_requests)} "
            f"slide(s) sequentially (one createSlide per batchUpdate call)."
        )
        await _execute_slides_sequentially(
            slides_service,
            presentation_id,
            creation_requests,
            label_prefix="slides.createSlide",
        )

        # Verification pass: refetch the deck and confirm every pre-allocated
        # placeholder objectId actually materialized. Custom layouts (esp.
        # two-column / multi-body) sometimes silently drop a placeholder
        # mapping at createSlide time — the API returns 200 and a slide is
        # created, but the placeholder we asked for is never bound. Any
        # downstream insertText / replaceImage targeting that ghost objectId
        # then triggers a generic HTTP 500 ("Internal error encountered")
        # for the entire batchUpdate. Filtering those orphans here turns a
        # hard failure into a logged warning while keeping the rest of the
        # slide intact.
        verify_pres = await _run_with_transient_retry(
            slides_service.presentations().get(presentationId=presentation_id).execute,
            label="slides.presentations.get (placeholder verification)",
        )
        existing_ids = _collect_existing_object_ids(verify_pres)
        # Phase B will create more objects (tables, free shapes, images,
        # sheets charts) with objectIds we minted client-side. Their
        # dependent requests (e.g. cell `insertText` into a newly created
        # table) target these IDs and would be wrongly filtered if we only
        # consulted the live presentation. Pre-register every objectId
        # that any queued create* request is about to mint.
        will_be_created: set = set()
        for _, reqs in per_slide_content:
            for req in reqs:
                created_id = _extract_created_object_id(req)
                if created_id:
                    will_be_created.add(created_id)
        existing_ids |= will_be_created
        # Also remember the actual placeholders per slide so we can suggest a
        # remediation in the warning (which type/index DID materialize).
        verify_slide_placeholders: Dict[str, List[Dict[str, Any]]] = {}
        for page in verify_pres.get("slides", []) or []:
            page_id = page.get("objectId")
            if not page_id:
                continue
            phs: List[Dict[str, Any]] = []
            for el in page.get("pageElements", []) or []:
                placeholder = (el.get("shape") or {}).get("placeholder") or {}
                if not placeholder:
                    continue
                phs.append(
                    {
                        "type": placeholder.get("type"),
                        "index": placeholder.get("index", 0),
                        "objectId": el.get("objectId"),
                    }
                )
            verify_slide_placeholders[page_id] = phs

        # Resolve deferred placeholder lookups (multi-same-type bindings).
        # For Two Columns and other custom layouts with 2+ placeholders of
        # the same type, we deliberately did NOT use placeholderIdMappings
        # at createSlide time (it triggers a Slides backend bug where the
        # second+ binding is a "ghost" — objectId visible in pageElements
        # but underlying shape unwired, causing HTTP 500 on insertText).
        # Now that the slides exist with auto-assigned IDs, we look up the
        # real objectId for each (slide_id, ph_type, layout_index) we
        # registered in `all_deferred_lookups` and rewrite every content
        # request that targets the corresponding pseudo objectId.
        rebind_map: Dict[str, str] = {}
        unresolved_lookups: List[str] = []
        for pseudo_id, (slide_id_for_lookup, ph_type, layout_index) in all_deferred_lookups.items():
            phs = verify_slide_placeholders.get(slide_id_for_lookup, [])
            real_id: Optional[str] = None
            for ph in phs:
                if ph.get("type") == ph_type and ph.get("index", 0) == layout_index:
                    real_id = ph.get("objectId")
                    break
            if real_id:
                rebind_map[pseudo_id] = real_id
            else:
                unresolved_lookups.append(
                    f"{pseudo_id}->({ph_type},index={layout_index},slide={slide_id_for_lookup})"
                )

        if unresolved_lookups:
            logger.warning(
                f"[create_audit_presentation] {len(unresolved_lookups)} deferred "
                f"placeholder lookup(s) could not be resolved on the live deck "
                f"(layout placeholder did not materialize as expected): "
                f"{unresolved_lookups[:5]}{' ...' if len(unresolved_lookups) > 5 else ''}. "
                f"Their content requests will be dropped by the orphan filter."
            )

        if rebind_map:
            logger.info(
                f"[create_audit_presentation] Rebinding {len(rebind_map)} deferred "
                f"placeholder ID(s) to their real auto-assigned objectIds (workaround "
                f"for Slides multi-same-type-placeholder ghost-binding bug)."
            )
            rewritten_per_slide_content: List[Tuple[int, List[Dict[str, Any]]]] = []
            for slide_idx, reqs in per_slide_content:
                rewritten_reqs: List[Dict[str, Any]] = []
                for req in reqs:
                    rewritten_reqs.append(_rebind_request_object_id(req, rebind_map))
                rewritten_per_slide_content.append((slide_idx, rewritten_reqs))
            per_slide_content = rewritten_per_slide_content

        filtered_per_slide_content: List[Tuple[int, List[Dict[str, Any]]]] = []
        dropped_total = 0
        for slide_idx, reqs in per_slide_content:
            kept: List[Dict[str, Any]] = []
            slide_id_for_log = slide_id_index_pairs[slide_idx][0] if slide_idx < len(
                slide_id_index_pairs
            ) else None
            for req in reqs:
                target_id = _extract_target_object_id(req)
                if target_id and target_id not in existing_ids:
                    dropped_total += 1
                    actual_phs = verify_slide_placeholders.get(slide_id_for_log, [])
                    logger.warning(
                        f"[create_audit_presentation] Slide #{slide_idx + 1}: dropping "
                        f"{next(iter(req.keys()))} request — target objectId "
                        f"'{target_id}' did not materialize on the slide. The custom "
                        f"layout silently dropped its placeholderIdMapping (likely a "
                        f"type/index mismatch with the layout's real placeholders). "
                        f"Slide actually exposes: "
                        f"{json.dumps(actual_phs, ensure_ascii=False)}"
                    )
                    continue
                kept.append(req)
            filtered_per_slide_content.append((slide_idx, kept))
        if dropped_total:
            logger.warning(
                f"[create_audit_presentation] Verification pass dropped {dropped_total} "
                f"orphan content request(s). The deck will still be generated, but the "
                f"affected fields will be empty. Open the layout(s) cited above in the "
                f"Slides editor and add the missing placeholders to fix permanently."
            )
        per_slide_content = filtered_per_slide_content

        # Phase B: fill every slide's content. Requests are grouped PER SLIDE
        # so a single batchUpdate never spans multiple slides — this is what
        # eliminates the residual HTTP 500s on heterogeneous content batches
        # against custom-layout templates. Heavy slides (e.g. big tables) are
        # still chunked at MAX_REQUESTS_PER_BATCH within their own slide.
        total_content_requests = sum(len(reqs) for _, reqs in per_slide_content)
        logger.info(
            f"[create_audit_presentation] Phase B: applying {total_content_requests} "
            f"content request(s) per-slide (≤{MAX_REQUESTS_PER_BATCH} per batchUpdate, "
            f"never crossing slide boundaries)."
        )
        await _execute_slides_per_slide(
            slides_service, presentation_id, per_slide_content
        )

        # 8) Speaker notes pass: re-fetch to find each slide's speakerNotesObjectId, then insert.
        notes_specs: List[Tuple[int, str]] = [
            (i, s["speaker_notes"]) for i, s in enumerate(slides) if s.get("speaker_notes")
        ]
        if notes_specs:
            updated_presentation = await _run_with_transient_retry(
                slides_service.presentations().get(presentationId=presentation_id).execute,
                label="slides.presentations.get (notes pass)",
            )
            slide_objs = updated_presentation.get("slides", []) or []
            notes_requests: List[Dict[str, Any]] = []
            for i, notes_text in notes_specs:
                actual_index = slide_offset + i
                if actual_index >= len(slide_objs):
                    continue
                notes_page = (
                    slide_objs[actual_index].get("slideProperties", {}).get("notesPage") or {}
                )
                speaker_notes_id = (
                    notes_page.get("notesProperties", {}).get("speakerNotesObjectId")
                )
                if not speaker_notes_id:
                    logger.warning(
                        f"[create_audit_presentation] Slide #{i + 1} has no speakerNotesObjectId; "
                        "skipping notes."
                    )
                    continue
                notes_requests.extend(B.build_speaker_notes_requests(speaker_notes_id, notes_text))
            await _execute_slides_batches(slides_service, presentation_id, notes_requests)

        # 9) Optional cleanup of data sheet.
        if cleanup_data_sheet and data_sheet_meta:
            await _delete_file_safe(
                drive_service,
                data_sheet_meta["spreadsheet_id"],
                "data sheet (cleanup_data_sheet=True)",
            )
            data_sheet_meta = None

    except Exception:
        if keep_on_error:
            # Debug mode: leave the partial files in Drive so the user can
            # inspect exactly which slides / placeholders / objects made it.
            # The IDs are logged loudly so they're easy to grep for.
            logger.warning(
                f"[create_audit_presentation] keep_on_error=True — leaving partial "
                f"deck on Drive for debugging: presentation_id={presentation_id} "
                f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            )
            if data_sheet_meta and data_sheet_meta.get("spreadsheet_id"):
                logger.warning(
                    f"[create_audit_presentation] keep_on_error=True — leaving partial "
                    f"data sheet on Drive: spreadsheet_id="
                    f"{data_sheet_meta.get('spreadsheet_id')} "
                    f"https://docs.google.com/spreadsheets/d/"
                    f"{data_sheet_meta.get('spreadsheet_id')}/edit"
                )
        else:
            # Best-effort rollback.
            await _delete_file_safe(drive_service, presentation_id, "partial deck")
            if data_sheet_meta:
                await _delete_file_safe(
                    drive_service, data_sheet_meta.get("spreadsheet_id"), "partial data sheet"
                )
        raise

    message_parts = [
        f"Created audit presentation '{final_title}' with {len(slides)} slide(s) for {user_google_email}."
    ]
    if folder_path_summary:
        message_parts.append(f"Folder: {folder_path_summary}.")
    if data_sheet_meta:
        message_parts.append(
            f"Data sheet with {data_sheet_meta['chart_count']} chart(s): "
            f"{data_sheet_meta['spreadsheet_url']}."
        )

    result = {
        "presentation_id": presentation_id,
        "presentation_url": presentation_url,
        "slide_count": len(slides),
        "data_sheet_url": data_sheet_meta["spreadsheet_url"] if data_sheet_meta else None,
        "data_sheet_id": data_sheet_meta["spreadsheet_id"] if data_sheet_meta else None,
        "folder_id": target_folder_id,
        "folder_path": folder_path_summary or None,
        "title": final_title,
        "message": " ".join(message_parts),
    }

    logger.info(
        f"[create_audit_presentation] Success: deck {presentation_id} ({len(slides)} slides), "
        f"data_sheet={data_sheet_meta['spreadsheet_id'] if data_sheet_meta else 'none'}"
    )
    return json.dumps(result, indent=2)
