"""
Pure helpers that build Google Slides API request payloads for the audit deck generator.

No I/O, no service calls. All functions return lists of `requests` dicts ready to be sent in
`presentations.batchUpdate(body={"requests": [...]})`.

Coordinate convention: positions and sizes use points (PT). The Slides API accepts both EMU and PT
when `unit` is set explicitly on the magnitude objects.
"""
from __future__ import annotations

import uuid
from typing import Any, Dict, Iterable, List, Optional, Tuple

# Google Slides predefined layout names. Used as fallbacks and aliases.
PREDEFINED_LAYOUTS = {
    "BLANK",
    "CAPTION_ONLY",
    "TITLE",
    "TITLE_AND_BODY",
    "TITLE_AND_TWO_COLUMNS",
    "TITLE_ONLY",
    "SECTION_HEADER",
    "SECTION_TITLE_AND_DESCRIPTION",
    "ONE_COLUMN_TEXT",
    "MAIN_POINT",
    "BIG_NUMBER",
}

# Friendly aliases the workflow can use in the JSON schema.
LAYOUT_ALIASES = {
    "title": "TITLE",
    "title_and_body": "TITLE_AND_BODY",
    "blank": "BLANK",
    "section_header": "SECTION_HEADER",
    "two_columns": "TITLE_AND_TWO_COLUMNS",
    "big_number": "BIG_NUMBER",
}

DEFAULT_PAGE_W_PT = 720.0  # 10in standard widescreen
DEFAULT_PAGE_H_PT = 405.0  # 5.625in widescreen


def gen_id(prefix: str) -> str:
    """Generate a short, deterministic-ish object ID safe for the Slides API.

    Slides object IDs must be unique per presentation, max 50 chars, alphanumeric/underscore/dash.
    """
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def resolve_layout_reference(
    presentation: Dict[str, Any], layout_name: str
) -> Dict[str, Any]:
    """Build a `slideLayoutReference` for `createSlide`.

    Resolution order:
      1. Predefined layout name (e.g. "TITLE_AND_BODY") -> {predefinedLayout: ...}
      2. Custom layout DISPLAY name found in the template's masters -> {layoutId: ...}
      3. Falls back to BLANK if nothing matches.
    """
    if not layout_name:
        return {"predefinedLayout": "BLANK"}

    canonical = LAYOUT_ALIASES.get(layout_name.lower(), layout_name)

    if canonical in PREDEFINED_LAYOUTS:
        return {"predefinedLayout": canonical}

    for layout in presentation.get("layouts", []) or []:
        props = layout.get("layoutProperties", {}) or {}
        display_name = props.get("displayName") or props.get("name")
        if display_name == layout_name or display_name == canonical:
            return {"layoutId": layout["objectId"]}

    return {"predefinedLayout": "BLANK"}


def _pt(value: float) -> Dict[str, Any]:
    return {"magnitude": float(value), "unit": "PT"}


def _size(width_pt: float, height_pt: float) -> Dict[str, Any]:
    return {"width": _pt(width_pt), "height": _pt(height_pt)}


def _transform(x_pt: float, y_pt: float) -> Dict[str, Any]:
    return {
        "scaleX": 1,
        "scaleY": 1,
        "translateX": float(x_pt),
        "translateY": float(y_pt),
        "unit": "PT",
    }


def _position(spec: Optional[Dict[str, Any]], default: Dict[str, float]) -> Dict[str, float]:
    """Merge a user-supplied position dict with defaults, normalizing keys to floats."""
    out = dict(default)
    if spec:
        for k in ("x", "y", "w", "h"):
            if k in spec and spec[k] is not None:
                out[k] = float(spec[k])
    return out


def build_text_insert_requests(
    object_id: str, text: str, style: Optional[Dict[str, Any]] = None
) -> List[Dict[str, Any]]:
    """Insert text into a shape, optionally applying a text style to the whole inserted range."""
    if not text:
        return []
    requests: List[Dict[str, Any]] = [
        {"insertText": {"objectId": object_id, "insertionIndex": 0, "text": text}}
    ]
    if style:
        requests.append(
            {
                "updateTextStyle": {
                    "objectId": object_id,
                    "textRange": {"type": "ALL"},
                    "style": style,
                    "fields": ",".join(style.keys()),
                }
            }
        )
    return requests


def build_text_box(
    slide_id: str,
    text: str,
    position: Dict[str, float],
    style: Optional[Dict[str, Any]] = None,
    paragraph_alignment: Optional[str] = None,
) -> Tuple[str, List[Dict[str, Any]]]:
    """Create a free-floating TEXT_BOX shape with text. Returns (shape_id, requests)."""
    shape_id = gen_id("tb")
    requests: List[Dict[str, Any]] = [
        {
            "createShape": {
                "objectId": shape_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": _size(position["w"], position["h"]),
                    "transform": _transform(position["x"], position["y"]),
                },
            }
        }
    ]
    requests.extend(build_text_insert_requests(shape_id, text, style))
    if paragraph_alignment:
        requests.append(
            {
                "updateParagraphStyle": {
                    "objectId": shape_id,
                    "textRange": {"type": "ALL"},
                    "style": {"alignment": paragraph_alignment},
                    "fields": "alignment",
                }
            }
        )
    return shape_id, requests


def build_create_slide(
    slide_id: str,
    layout_reference: Dict[str, Any],
    insertion_index: Optional[int] = None,
    placeholder_id_mappings: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    req: Dict[str, Any] = {
        "createSlide": {
            "objectId": slide_id,
            "slideLayoutReference": layout_reference,
        }
    }
    if insertion_index is not None:
        req["createSlide"]["insertionIndex"] = insertion_index
    if placeholder_id_mappings:
        req["createSlide"]["placeholderIdMappings"] = placeholder_id_mappings
    return req


def build_table_requests(
    slide_id: str,
    table_spec: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """Create a table on a slide and fill its cells.

    `table_spec` shape:
      {
        "headers": ["Metric", "Value"],
        "rows": [["...", "..."], ...],
        "position": {"x": 50, "y": 100, "w": 600, "h": 300},  # PT, optional
        "header_style": {...},  # optional textStyle for the header row
        "body_style": {...},    # optional textStyle for body rows
      }
    """
    headers = table_spec.get("headers") or []
    body_rows = table_spec.get("rows") or []
    if not headers and not body_rows:
        return []

    if headers and body_rows:
        all_rows = [headers] + list(body_rows)
    elif headers:
        all_rows = [headers]
    else:
        all_rows = list(body_rows)

    n_rows = len(all_rows)
    n_cols = max(len(r) for r in all_rows) if all_rows else 1

    pos = _position(
        table_spec.get("position"),
        {"x": 40.0, "y": 90.0, "w": DEFAULT_PAGE_W_PT - 80.0, "h": DEFAULT_PAGE_H_PT - 130.0},
    )

    table_id = gen_id("tbl")
    requests: List[Dict[str, Any]] = [
        {
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": _size(pos["w"], pos["h"]),
                    "transform": _transform(pos["x"], pos["y"]),
                },
                "rows": n_rows,
                "columns": n_cols,
            }
        }
    ]

    header_style = table_spec.get("header_style") or {"bold": True}
    body_style = table_spec.get("body_style")

    for r_idx, row in enumerate(all_rows):
        for c_idx in range(n_cols):
            value = row[c_idx] if c_idx < len(row) else ""
            text = "" if value is None else str(value)
            if not text:
                continue
            requests.append(
                {
                    "insertText": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
                        "text": text,
                        "insertionIndex": 0,
                    }
                }
            )
            cell_style = header_style if (headers and r_idx == 0) else body_style
            if cell_style:
                requests.append(
                    {
                        "updateTextStyle": {
                            "objectId": table_id,
                            "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
                            "textRange": {"type": "ALL"},
                            "style": cell_style,
                            "fields": ",".join(cell_style.keys()),
                        }
                    }
                )

    return requests


def build_image_requests(
    slide_id: str,
    image_spec: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """Insert an image from a URL.

    `image_spec` shape:
      {
        "url": "https://...",
        "position": {"x": 50, "y": 100, "w": 600, "h": 300},  # PT, optional
      }
    """
    url = image_spec.get("url")
    if not url:
        return []
    pos = _position(
        image_spec.get("position"),
        {"x": 60.0, "y": 100.0, "w": DEFAULT_PAGE_W_PT - 120.0, "h": DEFAULT_PAGE_H_PT - 150.0},
    )
    image_id = gen_id("img")
    return [
        {
            "createImage": {
                "objectId": image_id,
                "url": url,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": _size(pos["w"], pos["h"]),
                    "transform": _transform(pos["x"], pos["y"]),
                },
            }
        }
    ]


def build_sheets_chart_requests(
    slide_id: str,
    spreadsheet_id: str,
    chart_id: int,
    position: Optional[Dict[str, float]] = None,
    linking_mode: str = "LINKED",
) -> List[Dict[str, Any]]:
    """Embed a Google Sheets chart on a slide via its chartId."""
    pos = _position(
        position,
        {"x": 60.0, "y": 100.0, "w": DEFAULT_PAGE_W_PT - 120.0, "h": DEFAULT_PAGE_H_PT - 150.0},
    )
    return [
        {
            "createSheetsChart": {
                "objectId": gen_id("ch"),
                "spreadsheetId": spreadsheet_id,
                "chartId": int(chart_id),
                "linkingMode": linking_mode,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": _size(pos["w"], pos["h"]),
                    "transform": _transform(pos["x"], pos["y"]),
                },
            }
        }
    ]


def build_speaker_notes_requests(
    speaker_notes_object_id: str, notes_text: str
) -> List[Dict[str, Any]]:
    """Replace any existing speaker notes text on the notes shape, then insert the new text."""
    if not notes_text:
        return []
    return [
        {
            "deleteText": {
                "objectId": speaker_notes_object_id,
                "textRange": {"type": "ALL"},
            }
        },
        {
            "insertText": {
                "objectId": speaker_notes_object_id,
                "insertionIndex": 0,
                "text": notes_text,
            }
        },
    ]


def chunk_requests(
    requests: Iterable[Dict[str, Any]], max_per_batch: int
) -> List[List[Dict[str, Any]]]:
    """Split a flat list of requests into chunks of at most `max_per_batch` items."""
    chunks: List[List[Dict[str, Any]]] = []
    current: List[Dict[str, Any]] = []
    for r in requests:
        current.append(r)
        if len(current) >= max_per_batch:
            chunks.append(current)
            current = []
    if current:
        chunks.append(current)
    return chunks


def build_slide_with_placeholders(
    presentation: Dict[str, Any],
    slide_spec: Dict[str, Any],
    insertion_index: Optional[int] = None,
) -> Tuple[str, List[Dict[str, Any]], Dict[str, str]]:
    """Build the createSlide + placeholder text-fill + extra-element requests for ONE slide.

    Returns:
        (slide_id, requests, placeholder_ids)
        placeholder_ids maps semantic names ("title", "subtitle", "body") to the
        objectIds we assigned via `placeholderIdMappings`.
    """
    slide_id = gen_id("sl")
    layout_name = slide_spec.get("layout") or "BLANK"
    layout_ref = resolve_layout_reference(presentation, layout_name)

    fields = slide_spec.get("fields") or {}
    placeholder_ids: Dict[str, str] = {}
    placeholder_mappings: List[Dict[str, Any]] = []

    # Single-instance text placeholders (always at layout index 0).
    simple_text_fields = {
        "title": "TITLE",
        "centered_title": "CENTERED_TITLE",
        "subtitle": "SUBTITLE",
    }
    for field_name, ph_type in simple_text_fields.items():
        if field_name in fields and fields[field_name]:
            ph_id = gen_id("ph")
            placeholder_ids[field_name] = ph_id
            placeholder_mappings.append(
                {
                    "layoutPlaceholder": {"type": ph_type, "index": 0},
                    "objectId": ph_id,
                }
            )

    # BODY: supports a string (single BODY at index 0) OR a list (multi-column
    # layouts where the layout exposes BODY at index 0, 1, 2, ...).
    body_value = fields.get("body")
    body_texts: List[str] = []
    if isinstance(body_value, list):
        body_texts = [("" if v is None else str(v)) for v in body_value]
    elif body_value:
        body_texts = [str(body_value)]
    for i, body_text in enumerate(body_texts):
        if not body_text:
            continue
        ph_id = gen_id("ph")
        placeholder_ids[f"body[{i}]"] = ph_id
        placeholder_mappings.append(
            {
                "layoutPlaceholder": {"type": "BODY", "index": i},
                "objectId": ph_id,
            }
        )

    # PICTURE placeholders ("espace réservé image"). Specified per slide as
    # `image_placeholders`: either a list of URL strings or a list of dicts
    # of the form {"url": "...", "method": "CENTER_INSIDE"|"CENTER_CROP"}.
    image_placeholder_specs = slide_spec.get("image_placeholders") or []
    image_fill: List[Tuple[str, Dict[str, Any]]] = []  # (placeholder_id, spec)
    for i, raw in enumerate(image_placeholder_specs):
        if isinstance(raw, str):
            spec = {"url": raw}
        elif isinstance(raw, dict):
            spec = raw
        else:
            continue
        if not spec.get("url"):
            continue
        ph_id = gen_id("ph")
        placeholder_ids[f"image[{i}]"] = ph_id
        placeholder_mappings.append(
            {
                "layoutPlaceholder": {"type": "PICTURE", "index": i},
                "objectId": ph_id,
            }
        )
        image_fill.append((ph_id, spec))

    requests: List[Dict[str, Any]] = [
        build_create_slide(
            slide_id=slide_id,
            layout_reference=layout_ref,
            insertion_index=insertion_index,
            placeholder_id_mappings=placeholder_mappings or None,
        )
    ]

    # Fill simple single-instance text placeholders.
    for field_name in simple_text_fields:
        ph_id = placeholder_ids.get(field_name)
        if not ph_id:
            continue
        text = str(fields.get(field_name) or "")
        style = (slide_spec.get("styles") or {}).get(field_name)
        requests.extend(build_text_insert_requests(ph_id, text, style))

    # Fill BODY placeholder(s). Style may be a single dict (applied to all
    # body shapes) or a list aligned with the body texts.
    body_style = (slide_spec.get("styles") or {}).get("body")
    for i, body_text in enumerate(body_texts):
        ph_id = placeholder_ids.get(f"body[{i}]")
        if not ph_id:
            continue
        if isinstance(body_style, list):
            style = body_style[i] if i < len(body_style) else None
        else:
            style = body_style
        requests.extend(build_text_insert_requests(ph_id, body_text, style))

    # Fill PICTURE placeholder(s) via replaceImage. The placeholder we mapped
    # is created on the slide as an Image element holding the layout's
    # placeholder image; replaceImage swaps its bitmap for our URL while
    # preserving the placeholder's size, position, and crop.
    for ph_id, spec in image_fill:
        method = spec.get("method") or "CENTER_INSIDE"
        requests.append(
            {
                "replaceImage": {
                    "imageObjectId": ph_id,
                    "url": spec["url"],
                    "imageReplaceMethod": method,
                }
            }
        )

    # Free-floating title for BLANK-ish layouts when caller passes top-level "title".
    standalone_title = slide_spec.get("title")
    if standalone_title and "title" not in placeholder_ids:
        _, title_requests = build_text_box(
            slide_id=slide_id,
            text=str(standalone_title),
            position={"x": 40.0, "y": 30.0, "w": DEFAULT_PAGE_W_PT - 80.0, "h": 50.0},
            style={"bold": True, "fontSize": {"magnitude": 22, "unit": "PT"}},
        )
        requests.extend(title_requests)

    if "table" in slide_spec and slide_spec["table"]:
        requests.extend(build_table_requests(slide_id, slide_spec["table"]))

    if "image" in slide_spec and slide_spec["image"]:
        requests.extend(build_image_requests(slide_id, slide_spec["image"]))

    if "text_boxes" in slide_spec and slide_spec["text_boxes"]:
        for tb in slide_spec["text_boxes"]:
            text = tb.get("text", "")
            position = _position(
                tb.get("position"),
                {"x": 40.0, "y": 100.0, "w": 300.0, "h": 100.0},
            )
            _, tb_requests = build_text_box(
                slide_id=slide_id,
                text=text,
                position=position,
                style=tb.get("style"),
                paragraph_alignment=tb.get("alignment"),
            )
            requests.extend(tb_requests)

    return slide_id, requests, placeholder_ids
