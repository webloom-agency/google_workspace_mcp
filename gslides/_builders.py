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


def _normalize_layout_name(name: str) -> str:
    """Lowercase + collapse internal whitespace so 'Title  +  Body' matches 'title + body'."""
    return " ".join(str(name).strip().lower().split())


def _list_template_layouts(presentation: Dict[str, Any]) -> List[str]:
    """Return display names of every custom layout in the copied template."""
    names: List[str] = []
    for layout in presentation.get("layouts", []) or []:
        props = layout.get("layoutProperties", {}) or {}
        name = props.get("displayName") or props.get("name")
        if name:
            names.append(name)
    return names


def get_element_placeholder(element: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Return the placeholder dict on a PageElement, regardless of element kind.

    The Slides API exposes placeholders on the kind-specific subobject of each
    PageElement, NOT at the top level. Most placeholders ride on `shape`
    (TITLE, BODY, SUBTITLE, SLIDE_NUMBER, ...), but PICTURE placeholders
    created via the editor's `Insert → Placeholder → Image → Rectangle` path
    are typed as `image` PageElements with the placeholder field hanging off
    `image.placeholder`. Inspecting only `shape.placeholder` makes those
    PICTURE slots invisible to the rest of the build pipeline — the layout
    diagnostic claims the slot doesn't exist, `placeholderIdMappings` for
    `image_placeholders` find no target, and `replaceImage` is dropped by the
    orphan filter.

    Returns the placeholder dict (with `type` and optional `index`) if the
    element carries one in either location, else None.
    """
    shape_ph = (element.get("shape") or {}).get("placeholder")
    if shape_ph:
        return shape_ph
    image_ph = (element.get("image") or {}).get("placeholder")
    if image_ph:
        return image_ph
    return None


def get_layout_placeholders_by_type(
    presentation: Dict[str, Any], layout_object_id: str
) -> Dict[str, List[int]]:
    """Return the *actual* placeholder indexes a layout exposes, grouped by type.

    Custom layouts created in the Slides editor (and derived layouts whose
    internal `name` looks like 'TITLE_AND_BODY_1_2') do NOT necessarily have
    placeholders at sequential indexes 0, 1, 2... Google assigns indexes at
    creation time and they can be sparse / non-zero. Hard-coding `index: 0`
    in `placeholderIdMappings` then either silently fails to bind (and our
    insertText hits a non-existent objectId) or — more commonly — makes the
    Slides backend return a non-specific HTTP 500 ('Internal error
    encountered') for the entire `createSlide` request.

    Use this helper to discover the real indexes for the layout the slide
    targets, then map each requested placeholder to one of them.

    Returns a dict like {"TITLE": [0], "BODY": [3, 7], "PICTURE": [9]} where
    every list is sorted ascending.
    """
    out: Dict[str, List[int]] = {}
    for layout in presentation.get("layouts", []) or []:
        if layout.get("objectId") != layout_object_id:
            continue
        for element in layout.get("pageElements", []) or []:
            placeholder = get_element_placeholder(element)
            if not placeholder:
                continue
            ptype = placeholder.get("type")
            pindex = placeholder.get("index", 0)
            if ptype:
                out.setdefault(ptype, []).append(int(pindex))
        break
    for ptype in out:
        out[ptype] = sorted(set(out[ptype]))
    return out


def get_layout_placeholder_geometry(
    presentation: Dict[str, Any],
    layout_object_id: str,
    ph_type: str,
    occurrence: int = 0,
) -> Optional[Tuple[Dict[str, Any], Dict[str, Any]]]:
    """Return (size, transform) of the `occurrence`-th `ph_type` placeholder
    on the layout `layout_object_id`, in the exact format the Slides API
    `Size` / `AffineTransform` use.

    `occurrence` orders by ascending `placeholder.index` (so occurrence=0 is
    the placeholder with the lowest index for that type — typically left or
    top).

    Returns None if the layout is not found or has no such placeholder.

    This is used to overlay a free-floating TEXT_BOX on top of broken
    multi-same-type placeholders (Slides API ghost-bug workaround).
    """
    for layout in presentation.get("layouts", []) or []:
        if layout.get("objectId") != layout_object_id:
            continue
        candidates: List[Tuple[int, Dict[str, Any]]] = []
        for element in layout.get("pageElements", []) or []:
            placeholder = get_element_placeholder(element)
            if not placeholder or placeholder.get("type") != ph_type:
                continue
            candidates.append((int(placeholder.get("index", 0)), element))
        candidates.sort(key=lambda t: t[0])
        if occurrence >= len(candidates):
            return None
        _, el = candidates[occurrence]
        size = el.get("size")
        transform = el.get("transform")
        if not size or not transform:
            return None
        return size, transform
    return None


def resolve_layout_reference(
    presentation: Dict[str, Any], layout_name: str
) -> Dict[str, Any]:
    """Build a `slideLayoutReference` for `createSlide`.

    Resolution order:
      1. Predefined layout name (e.g. "TITLE_AND_BODY") -> {predefinedLayout: ...}
      2. Custom layout DISPLAY name found in the template's masters -> {layoutId: ...}
         Match is case-insensitive and whitespace-normalized so 'Title + Body',
         'title  +  body', and 'TITLE + BODY' all resolve identically.
      3. Raises a clear Exception listing every available custom layout. We do
         NOT fall back to predefined BLANK here because most user templates
         ship a custom master that doesn't include the BLANK predefined layout
         (Google then 400s with 'predefined layout (BLANK) is not present in
         the current master'), which masks the real misconfiguration.
    """
    if not layout_name:
        raise Exception(
            "Slide is missing 'layout'. Set it to a predefined name "
            "(e.g. 'TITLE_AND_BODY', 'SECTION_HEADER') or to the exact display "
            "name of a custom layout in your template."
        )

    # Predefined layout (or one of its friendly aliases).
    canonical = LAYOUT_ALIASES.get(layout_name.lower(), layout_name)
    if canonical in PREDEFINED_LAYOUTS:
        return {"predefinedLayout": canonical}

    # Custom layout — match by normalized display name.
    target = _normalize_layout_name(layout_name)
    canonical_target = _normalize_layout_name(canonical)
    for layout in presentation.get("layouts", []) or []:
        props = layout.get("layoutProperties", {}) or {}
        display_name = props.get("displayName") or props.get("name")
        if not display_name:
            continue
        if _normalize_layout_name(display_name) in (target, canonical_target):
            return {"layoutId": layout["objectId"]}

    # No match — surface the real problem instead of silently picking BLANK.
    available = _list_template_layouts(presentation)
    available_str = ", ".join(repr(n) for n in available) if available else "(none found)"
    raise Exception(
        f"Layout '{layout_name}' was not found in the template. "
        f"Available custom layouts in this template: {available_str}. "
        f"Predefined layouts you can also use: {sorted(PREDEFINED_LAYOUTS)}. "
        f"Either rename the slide's 'layout' to match one of the above exactly "
        f"(matching is case-insensitive and ignores extra spaces), or add a "
        f"custom layout with that display name to your template."
    )


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


def _utf16_len(s: str) -> int:
    """Length of `s` in UTF-16 code units (matches Google Slides textRange indexing).

    Emoji and other supplementary-plane chars contribute 2 units; BMP chars 1.
    """
    n = 0
    for ch in s:
        n += 2 if ord(ch) > 0xFFFF else 1
    return n


def _parse_inline_bold(text: str) -> Tuple[str, List[Tuple[int, int]]]:
    """Strip `**bold**` markers from `text`, return the plain string and a list
    of (start_utf16, end_utf16_exclusive) ranges that should be rendered bold.

    Indexes are in UTF-16 code units (Google Slides indexing convention) so
    emoji-heavy text bolds correctly.

    Edge cases:
      * Unmatched `**` (no closing pair) is left untouched.
      * `\\**` is treated as a literal `**` (escape hatch).
    """
    out_chars: List[str] = []
    bold_ranges: List[Tuple[int, int]] = []
    cursor16 = 0
    i = 0
    n = len(text)
    while i < n:
        if text.startswith("\\**", i):
            out_chars.append("**")
            cursor16 += 2
            i += 3
            continue
        if text.startswith("**", i):
            j = text.find("**", i + 2)
            if j != -1:
                inner = text[i + 2 : j]
                inner_len16 = _utf16_len(inner)
                if inner_len16 > 0:
                    bold_ranges.append((cursor16, cursor16 + inner_len16))
                out_chars.append(inner)
                cursor16 += inner_len16
                i = j + 2
                continue
        ch = text[i]
        out_chars.append(ch)
        cursor16 += 2 if ord(ch) > 0xFFFF else 1
        i += 1
    return "".join(out_chars), bold_ranges


def build_text_insert_requests(
    object_id: str, text: str, style: Optional[Dict[str, Any]] = None
) -> List[Dict[str, Any]]:
    """Insert text into a shape, optionally applying a text style to the whole inserted range.

    Inline `**bold**` markers in `text` are stripped and replaced with per-range
    `updateTextStyle` requests that bold the wrapped span. Use `\\**` to insert
    a literal pair of asterisks. Emojis are passed through as-is and correctly
    accounted for in the UTF-16 indexing the Slides API expects.
    """
    if not text:
        return []
    plain, bold_ranges = _parse_inline_bold(text)
    requests: List[Dict[str, Any]] = [
        {"insertText": {"objectId": object_id, "insertionIndex": 0, "text": plain}}
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
    for start, end in bold_ranges:
        if end <= start:
            continue
        requests.append(
            {
                "updateTextStyle": {
                    "objectId": object_id,
                    "textRange": {
                        "type": "FIXED_RANGE",
                        "startIndex": start,
                        "endIndex": end,
                    },
                    "style": {"bold": True},
                    "fields": "bold",
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
) -> Tuple[
    str,
    List[Dict[str, Any]],
    List[Dict[str, Any]],
    Dict[str, str],
    Dict[str, Tuple[str, str, int]],
]:
    """Build the createSlide + placeholder text-fill + extra-element requests for ONE slide.

    The returned requests are split into two phases so the caller can run *every*
    `createSlide` first (in its own batch) and *every* content mutation second.
    Mixing slide creation with dependent text inserts / table fills in a single
    Slides batchUpdate is the canonical trigger for HTTP 500s on larger decks
    (each createSlide forces placeholder ID materialization + layout inheritance,
    and stacking many of those next to dependent inserts in one batch is fragile).

    Returns:
        (slide_id, creation_requests, content_requests, placeholder_ids,
         deferred_placeholder_lookups)
        - creation_requests: the single `createSlide` request for this slide.
        - content_requests: insertText, updateTextStyle, createTable (+ cell
          inserts), createShape (text boxes), createImage, replaceImage. Every
          one of these references either the slide itself, a placeholder
          objectId we pre-allocated via `placeholderIdMappings`, or a *pseudo*
          objectId scheduled for post-Phase-A rebinding (see below).
        - placeholder_ids: maps semantic names ("title", "subtitle", "body[i]",
          "image[i]") to the objectIds (real or pseudo) used by the requests
          targeting that placeholder.
        - deferred_placeholder_lookups: maps pseudo objectIds → (slide_id,
          placeholder_type, occurrence). `occurrence` is the 0-based ORDER of
          this placeholder among same-type placeholders on the slide (matching
          by ascending slide-level `index`). The caller MUST resolve these to
          the actual auto-assigned objectIds (by reading the live deck after
          Phase A) and rewrite every content_request targeting them.

    Why deferred lookups exist:
        For custom layouts with MULTIPLE placeholders of the same type
        (e.g. two BODY placeholders in a "Two Columns" layout), the Slides
        backend has a known pathology: when `placeholderIdMappings` binds
        2+ same-type placeholders, the second+ binding ends up as a "ghost"
        — its objectId appears in the slide's pageElements (so verification
        passes) but the underlying shape isn't fully wired, so any operation
        on it returns a generic HTTP 500 ('Internal error encountered').
        For these cases we deliberately DO NOT include the placeholder in
        `placeholderIdMappings`; we let Slides auto-assign the objectId,
        then resolve our pseudo objectId to the real one after Phase A and
        rewrite the dependent content requests. This bypasses the bug.
    """
    slide_id = gen_id("sl")
    layout_name = slide_spec.get("layout") or "BLANK"
    layout_ref = resolve_layout_reference(presentation, layout_name)

    # Discover the layout's REAL placeholder indexes (only for custom layouts;
    # predefined layouts always expose unique placeholders at index 0). This
    # is critical because the Slides backend returns a non-specific HTTP 500
    # ('Internal error encountered') if a `placeholderIdMapping` references a
    # type/index combo the layout doesn't actually have — and custom layouts
    # in the editor get arbitrary indexes (BODY can land at 3, 7, etc).
    layout_placeholders_by_type: Dict[str, List[int]] = {}
    if "layoutId" in layout_ref:
        layout_placeholders_by_type = get_layout_placeholders_by_type(
            presentation, layout_ref["layoutId"]
        )

    def _is_multi_occurrence(ph_type: str) -> bool:
        """A custom layout has 2+ placeholders of `ph_type`?

        Predefined layouts never trigger the multi-mapping bug, so we only
        defer for custom layouts (`layoutId` is set).
        """
        if "layoutId" not in layout_ref:
            return False
        return len(layout_placeholders_by_type.get(ph_type) or []) > 1

    def _build_layout_placeholder_mapping(
        ph_type: str, occurrence: int = 0
    ) -> Optional[Dict[str, Any]]:
        """Build a `layoutPlaceholder` mapping object for the `occurrence`-th placeholder
        of `ph_type` in the resolved layout. Returns None if the layout has no such
        placeholder.

        Per Google's docs: when the layout has a single placeholder of a given type,
        OMITTING `index` is the canonical way to bind to it — the Slides backend
        will match by type alone. We only emit `index` when there are multiple
        placeholders of the same type (e.g. BODY[0] / BODY[1] in a two-column
        layout) where disambiguation is genuinely needed.

        For predefined layouts (no `layoutId` in `layout_ref`) we don't have the
        placeholder list, so we fall back to the legacy index = `occurrence` behavior,
        which works because predefined layouts are well-defined.
        """
        if "predefinedLayout" in layout_ref:
            mapping = {"type": ph_type}
            if occurrence > 0:
                mapping["index"] = occurrence
            return mapping
        indexes = layout_placeholders_by_type.get(ph_type) or []
        if occurrence >= len(indexes):
            return None
        if len(indexes) == 1:
            return {"type": ph_type}
        return {"type": ph_type, "index": indexes[occurrence]}

    fields = slide_spec.get("fields") or {}
    placeholder_ids: Dict[str, str] = {}
    placeholder_mappings: List[Dict[str, Any]] = []
    skipped_fields: List[str] = []
    # pseudo_id -> (slide_id, placeholder_type, occurrence_to_match)
    # `occurrence_to_match` is the 0-based ORDER of the placeholder among
    # same-type placeholders on the slide (NOT the layout-level `index` value).
    # We store occurrence (not the layout's raw `index` value) because
    # slide-level `placeholder.index` reported by `presentations.get` is not
    # always equal to the corresponding layout-level `index`. Matching by
    # ascending-`index` order is the only reliable way to map the N-th BODY of
    # the layout to the N-th BODY actually materialized on the slide.
    deferred_lookups: Dict[str, Tuple[str, str, int]] = {}

    def _allocate_placeholder(
        ph_type: str, occurrence: int, semantic_name: str
    ) -> Optional[str]:
        """Allocate either a pre-bound or deferred objectId for one placeholder.

        Returns the objectId to use in dependent requests, or None if the layout
        has no matching placeholder for (ph_type, occurrence).

        Strategy:
          * Singleton custom-layout placeholders (only one of `ph_type` in the
            layout) and predefined-layout placeholders → pre-bind via
            `placeholderIdMappings`. This works reliably.
          * Multi-occurrence custom-layout placeholders (BODY[0] AND BODY[1]
            in Two Columns, etc.) → emit a pseudo objectId now and defer real
            ID resolution until after Phase A. Avoids the multi-mapping ghost
            bug.
        """
        indexes = layout_placeholders_by_type.get(ph_type) or []
        if "layoutId" in layout_ref and occurrence >= len(indexes):
            return None
        ph_id = gen_id("ph")
        placeholder_ids[semantic_name] = ph_id
        if _is_multi_occurrence(ph_type):
            deferred_lookups[ph_id] = (slide_id, ph_type, occurrence)
            return ph_id
        # Singleton or predefined: traditional pre-binding works fine.
        mapping = _build_layout_placeholder_mapping(ph_type, occurrence=occurrence)
        if mapping is None:
            placeholder_ids.pop(semantic_name, None)
            return None
        placeholder_mappings.append({"layoutPlaceholder": mapping, "objectId": ph_id})
        return ph_id

    simple_text_fields = {
        "title": "TITLE",
        "centered_title": "CENTERED_TITLE",
        "subtitle": "SUBTITLE",
    }
    for field_name, ph_type in simple_text_fields.items():
        if field_name in fields and fields[field_name]:
            allocated = _allocate_placeholder(ph_type, 0, field_name)
            if allocated is None:
                skipped_fields.append(f"{field_name} ({ph_type})")

    body_value = fields.get("body")
    body_texts: List[str] = []
    if isinstance(body_value, list):
        body_texts = [("" if v is None else str(v)) for v in body_value]
    elif body_value:
        body_texts = [str(body_value)]

    # body indexes that should be rendered as a free-floating TEXT_BOX
    # overlay (workaround for the Slides multi-BODY ghost bug). We track
    # i -> (overlay_object_id, size, transform) so the content-building
    # phase below can emit the correct createShape + insertText sequence.
    body_overlays: Dict[int, Tuple[str, Dict[str, Any], Dict[str, Any]]] = {}

    for i, body_text in enumerate(body_texts):
        if not body_text:
            continue

        # Multi-occurrence custom-layout BODY placeholders: the FIRST one
        # (occurrence 0) accepts text via the deferred-rebind path. Every
        # subsequent BODY placeholder of the same layout is created in a
        # corrupt "ghost" state by Slides — `insertText` against it always
        # returns HTTP 500 regardless of how the objectId was bound. Bypass
        # the broken placeholder by laying a free-floating TEXT_BOX shape
        # over its layout-defined geometry. The slide-level placeholder
        # remains underneath but is empty, so its prompt is hidden behind
        # our text box (and prompts never render in present mode).
        if (
            "layoutId" in layout_ref
            and i >= 1
            and len(layout_placeholders_by_type.get("BODY") or []) > 1
        ):
            geom = get_layout_placeholder_geometry(
                presentation, layout_ref["layoutId"], "BODY", i
            )
            if geom is not None:
                size, transform = geom
                overlay_id = gen_id("body_tb")
                placeholder_ids[f"body[{i}]"] = overlay_id
                body_overlays[i] = (overlay_id, size, transform)
                continue
            # Geometry not available — fall through to the normal placeholder
            # path so we at least try (and skip cleanly if it fails).

        allocated = _allocate_placeholder("BODY", i, f"body[{i}]")
        if allocated is None:
            skipped_fields.append(f"body[{i}] (BODY)")

    image_placeholder_specs = slide_spec.get("image_placeholders") or []
    image_fill: List[Tuple[str, Dict[str, Any]]] = []
    for i, raw in enumerate(image_placeholder_specs):
        if isinstance(raw, str):
            spec = {"url": raw}
        elif isinstance(raw, dict):
            spec = raw
        else:
            continue
        if not spec.get("url"):
            continue
        allocated = _allocate_placeholder("PICTURE", i, f"image[{i}]")
        if allocated is None:
            skipped_fields.append(f"image[{i}] (PICTURE)")
            continue
        image_fill.append((allocated, spec))

    if skipped_fields:
        # Surface this as part of the returned placeholder_ids so the caller can
        # log it. We don't raise: a missing placeholder in the layout is a soft
        # mismatch, not a fatal error — better to render the rest of the slide
        # than to abort the whole deck.
        placeholder_ids["__skipped__"] = ",".join(skipped_fields)

    creation_requests: List[Dict[str, Any]] = [
        build_create_slide(
            slide_id=slide_id,
            layout_reference=layout_ref,
            insertion_index=insertion_index,
            placeholder_id_mappings=placeholder_mappings or None,
        )
    ]
    content_requests: List[Dict[str, Any]] = []

    # Fill simple single-instance text placeholders.
    for field_name in simple_text_fields:
        ph_id = placeholder_ids.get(field_name)
        if not ph_id:
            continue
        text = str(fields.get(field_name) or "")
        style = (slide_spec.get("styles") or {}).get(field_name)
        content_requests.extend(build_text_insert_requests(ph_id, text, style))

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

        overlay = body_overlays.get(i)
        if overlay is not None:
            # Multi-BODY ghost-bug workaround: emit a TEXT_BOX overlay at
            # the layout's body[i] geometry, then insertText into it. The
            # underlying broken slide-level placeholder is left alone but
            # is empty; our overlay covers its prompt visually.
            overlay_id, size, transform = overlay
            content_requests.append(
                {
                    "createShape": {
                        "objectId": overlay_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": size,
                            "transform": transform,
                        },
                    }
                }
            )
            content_requests.extend(
                build_text_insert_requests(overlay_id, body_text, style)
            )
            continue

        content_requests.extend(build_text_insert_requests(ph_id, body_text, style))

    # Fill PICTURE placeholder(s) via replaceImage. The placeholder we mapped
    # is created on the slide as an Image element holding the layout's
    # placeholder image; replaceImage swaps its bitmap for our URL while
    # preserving the placeholder's size, position, and crop.
    for ph_id, spec in image_fill:
        method = spec.get("method") or "CENTER_INSIDE"
        content_requests.append(
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
        content_requests.extend(title_requests)

    if "table" in slide_spec and slide_spec["table"]:
        content_requests.extend(build_table_requests(slide_id, slide_spec["table"]))

    if "image" in slide_spec and slide_spec["image"]:
        content_requests.extend(build_image_requests(slide_id, slide_spec["image"]))

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
            content_requests.extend(tb_requests)

    return slide_id, creation_requests, content_requests, placeholder_ids, deferred_lookups
