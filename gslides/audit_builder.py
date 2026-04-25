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
) -> Dict[str, Any]:
    """Build a single `addChart` request given a chart spec and the data range geometry."""
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

    if chart_type in _BASIC_CHART_TYPES:
        axis: List[Dict[str, Any]] = []
        if chart_spec.get("domain_axis_title"):
            axis.append({"position": "BOTTOM_AXIS", "title": chart_spec["domain_axis_title"]})
        if chart_spec.get("value_axis_title"):
            axis.append({"position": "LEFT_AXIS", "title": chart_spec["value_axis_title"]})

        spec["basicChart"] = {
            "chartType": chart_type,
            "legendPosition": chart_spec.get("legend_position", "BOTTOM_LEGEND"),
            "headerCount": 1,
            "domains": [{"domain": {"sourceRange": {"sources": [domain_source]}}}],
            "series": [
                {"series": {"sourceRange": {"sources": [src]}}, "targetAxis": "LEFT_AXIS"}
                for src in series_sources
            ],
        }
        if axis:
            spec["basicChart"]["axis"] = axis
    elif chart_type in _PIE_CHART_TYPES:
        spec["pieChart"] = {
            "legendPosition": chart_spec.get("legend_position", "BOTTOM_LEGEND"),
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
        "slides": [
          {"layout": "TITLE", "fields": {"title": "...", "subtitle": "..."}},
          {"layout": "TITLE_AND_BODY", "fields": {"title": "...", "body": "..."},
           "speaker_notes": "..."},
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

    Returns:
        str: JSON string with presentation_id, presentation_url, slide_count, data_sheet_url,
        folder_id, message.
    """
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

        # 5) Strip whatever slides the template ships with so we have a clean canvas.
        await _strip_existing_slides(slides_service, presentation_id)

        # 6) If any slide needs a chart, build the data Sheet first.
        flat_chart_specs = _annotate_chart_uids(deck)
        chart_id_by_uid: Dict[str, int] = {}
        if flat_chart_specs:
            data_sheet_meta, chart_pairs = await _create_data_sheet_and_charts(
                sheets_service,
                drive_service,
                title=final_title,
                target_folder_id=target_folder_id,
                chart_specs=flat_chart_specs,
            )
            chart_id_by_uid = dict(chart_pairs)

        # 7) Build all slide-creation requests in one big list (we'll chunk for execution).
        presentation = await asyncio.to_thread(
            slides_service.presentations().get(presentationId=presentation_id).execute
        )

        all_requests: List[Dict[str, Any]] = []
        slide_id_index_pairs: List[Tuple[str, int]] = []

        for index, slide_spec in enumerate(slides):
            slide_id, slide_requests, _placeholders = B.build_slide_with_placeholders(
                presentation=presentation,
                slide_spec=slide_spec,
                insertion_index=index,
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
                if i >= len(slide_objs):
                    continue
                notes_page = slide_objs[i].get("slideProperties", {}).get("notesPage") or {}
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
