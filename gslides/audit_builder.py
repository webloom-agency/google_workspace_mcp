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

from auth.service_decorator import require_multiple_services
from core.server import server
from core.utils import handle_http_errors

from gslides import _builders as B

logger = logging.getLogger(__name__)

# Google's `slides.batchUpdate` becomes flaky beyond ~500 requests / 2 MB. 400 leaves headroom.
MAX_REQUESTS_PER_BATCH = 400
# Per-call addChart batches stay comfortably under Sheets quota.
MAX_CHARTS_PER_SHEETS_BATCH = 20

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


# -----------------------------
# Drive / template helpers
# -----------------------------
async def _copy_template(
    drive_service, template_presentation_id: str, new_title: str
) -> Dict[str, str]:
    """Copy a template Slides file. Returns the new file's metadata."""
    body = {"name": new_title}
    return await asyncio.to_thread(
        drive_service.files()
        .copy(
            fileId=template_presentation_id,
            body=body,
            fields="id, name, parents, webViewLink",
            supportsAllDrives=True,
        )
        .execute
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
        axis: List[Dict[str, Any]] = []
        if chart_spec.get("domain_axis_title"):
            axis.append({"position": "BOTTOM_AXIS", "title": chart_spec["domain_axis_title"]})
        if chart_spec.get("value_axis_title"):
            axis.append({"position": "LEFT_AXIS", "title": chart_spec["value_axis_title"]})

        series_palette = style.get("series_colors") or []
        series_list: List[Dict[str, Any]] = []
        for idx, src in enumerate(series_sources):
            series_entry: Dict[str, Any] = {
                "series": {"sourceRange": {"sources": [src]}},
                "targetAxis": "LEFT_AXIS",
            }
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
    spreadsheet = await asyncio.to_thread(
        sheets_service.spreadsheets().create(body=body).execute
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

    await asyncio.to_thread(
        sheets_service.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": "USER_ENTERED", "data": value_ranges},
        )
        .execute
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
    for chunk in chunks:
        response = await asyncio.to_thread(
            sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": chunk})
            .execute
        )
        replies = response.get("replies", [])
        for j, reply in enumerate(replies):
            if "addChart" not in reply:
                raise Exception("Sheets API returned an unexpected reply (missing addChart).")
            chart_uid = chart_specs[chunk_offset + j].get("_uid", f"chart_{chunk_offset + j}")
            chart_ids.append((chart_uid, reply["addChart"]["chart"]["chartId"]))
        chunk_offset += len(chunk)

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
    presentation = await asyncio.to_thread(
        slides_service.presentations().get(presentationId=presentation_id).execute
    )
    slides = presentation.get("slides", []) or []
    if not slides:
        return
    requests = [{"deleteObject": {"objectId": s["objectId"]}} for s in slides]
    await asyncio.to_thread(
        slides_service.presentations()
        .batchUpdate(presentationId=presentation_id, body={"requests": requests})
        .execute
    )


async def _execute_slides_batches(
    slides_service, presentation_id: str, requests: List[Dict[str, Any]]
) -> None:
    """Run a flat list of Slides requests in chunks of MAX_REQUESTS_PER_BATCH."""
    if not requests:
        return
    for chunk in B.chunk_requests(requests, MAX_REQUESTS_PER_BATCH):
        await asyncio.to_thread(
            slides_service.presentations()
            .batchUpdate(presentationId=presentation_id, body={"requests": chunk})
            .execute
        )


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
        presentation = await asyncio.to_thread(
            slides_service.presentations().get(presentationId=presentation_id).execute
        )

        # When keeping template slides, append generated slides at the end so the boilerplate
        # (cover, methodology, ...) stays in front. Otherwise the template is already empty.
        slide_offset = len(presentation.get("slides", []) or []) if keep_template_slides else 0

        all_requests: List[Dict[str, Any]] = []
        slide_id_index_pairs: List[Tuple[str, int]] = []

        for index, slide_spec in enumerate(slides):
            slide_id, slide_requests, _placeholders = B.build_slide_with_placeholders(
                presentation=presentation,
                slide_spec=slide_spec,
                insertion_index=slide_offset + index,
            )
            slide_id_index_pairs.append((slide_id, index))

            chart_spec = slide_spec.get("chart")
            if chart_spec:
                uid = chart_spec.get("_uid")
                chart_id = chart_id_by_uid.get(uid)
                if chart_id is None:
                    raise Exception(
                        f"Internal error: chart for slide #{index + 1} was not created."
                    )
                slide_requests.extend(
                    B.build_sheets_chart_requests(
                        slide_id=slide_id,
                        spreadsheet_id=data_sheet_meta["spreadsheet_id"],
                        chart_id=chart_id,
                        position=chart_spec.get("position"),
                    )
                )

            all_requests.extend(slide_requests)

        await _execute_slides_batches(slides_service, presentation_id, all_requests)

        # 8) Speaker notes pass: re-fetch to find each slide's speakerNotesObjectId, then insert.
        notes_specs: List[Tuple[int, str]] = [
            (i, s["speaker_notes"]) for i, s in enumerate(slides) if s.get("speaker_notes")
        ]
        if notes_specs:
            updated_presentation = await asyncio.to_thread(
                slides_service.presentations().get(presentationId=presentation_id).execute
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
