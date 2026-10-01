"""
Pure helpers that build Google Slides API request payloads for the audit deck generator.

No I/O, no service calls. All functions return lists of `requests` dicts ready to be sent in
`presentations.batchUpdate(body={"requests": [...]})`.

Coordinate convention: positions and sizes use points (PT). The Slides API accepts both EMU and PT
when `unit` is set explicitly on the magnitude objects.
"""
from __future__ import annotations

import logging
import re
import uuid
from typing import Any, Dict, Iterable, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Slides ReplaceImageRequest.imageReplaceMethod — only these are valid.
_VALID_IMAGE_REPLACE_METHODS = frozenset(
    {
        "IMAGE_REPLACE_METHOD_UNSPECIFIED",
        "CENTER_INSIDE",
        "CENTER_CROP",
    }
)

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


def find_predefined_layout_id(
    presentation: Dict[str, Any], predefined_name: str
) -> Optional[str]:
    """Return the `objectId` of the layout whose internal name matches
    `predefined_name` (e.g. 'TITLE_AND_BODY'), or None if absent.

    Predefined Google layouts are exposed in `presentation.layouts[]` with
    `layoutProperties.name == "<PREDEFINED_NAME>"`. When a user customizes a
    template (deletes a placeholder from `TITLE_AND_BODY`, reorganizes
    placeholders in `TITLE_AND_TWO_COLUMNS`, etc.) the predefined layout
    keeps its name but its placeholder indexes can drift. Looking up the
    layoutId lets us discover the layout's REAL placeholder structure and
    avoid the same `placeholderIdMappings` trap we already work around for
    custom layouts (HTTP 400 'placeholder is not on the page').
    """
    if not predefined_name:
        return None
    target = predefined_name.upper()
    for layout in presentation.get("layouts", []) or []:
        props = layout.get("layoutProperties", {}) or {}
        name = (props.get("name") or "").upper()
        if name == target:
            return layout.get("objectId")
    return None


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


def _dimension_to_pt(dim: Optional[Dict[str, Any]]) -> float:
    """Convert a Slides Dimension ({magnitude, unit}) to points."""
    if not dim:
        return 0.0
    mag = float(dim.get("magnitude") or 0)
    unit = (dim.get("unit") or "EMU").upper()
    if unit == "PT":
        return mag
    # EMU (English Metric Unit): 12700 EMU = 1 PT
    return mag / 12700.0


def geometry_to_position(
    size: Dict[str, Any], transform: Dict[str, Any]
) -> Dict[str, float]:
    """Convert layout placeholder size+transform into {x,y,w,h} in PT."""
    scale_x = float(transform.get("scaleX") or 1.0)
    scale_y = float(transform.get("scaleY") or 1.0)
    w = _dimension_to_pt(size.get("width")) * abs(scale_x)
    h = _dimension_to_pt(size.get("height")) * abs(scale_y)
    # translate may be EMU or PT depending on transform.unit
    t_unit = (transform.get("unit") or "EMU").upper()
    tx = float(transform.get("translateX") or 0)
    ty = float(transform.get("translateY") or 0)
    if t_unit != "PT":
        tx /= 12700.0
        ty /= 12700.0
    return {"x": tx, "y": ty, "w": w, "h": h}


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


def _normalize_image_replace_method(raw: Any) -> str:
    """Map user/LLM input to a valid Slides `imageReplaceMethod` enum string."""
    if raw is None or raw == "":
        return "CENTER_INSIDE"
    s = str(raw).strip().upper().replace("-", "_")
    # Common shorthand from agents / UI copy.
    aliases = {
        "CENTER_INSIDE": "CENTER_INSIDE",
        "CENTER_CROP": "CENTER_CROP",
        "CROP": "CENTER_CROP",
        "FIT": "CENTER_INSIDE",
        "CONTAIN": "CENTER_INSIDE",
        "COVER": "CENTER_CROP",
        "IMAGE_REPLACE_METHOD_UNSPECIFIED": "IMAGE_REPLACE_METHOD_UNSPECIFIED",
        "UNSPECIFIED": "IMAGE_REPLACE_METHOD_UNSPECIFIED",
    }
    if s in aliases:
        return aliases[s]
    if s in _VALID_IMAGE_REPLACE_METHODS:
        return s
    logger.warning(
        f"[_builders] Invalid imageReplaceMethod {raw!r} — falling back to CENTER_INSIDE. "
        f"Valid: CENTER_INSIDE, CENTER_CROP."
    )
    return "CENTER_INSIDE"


def _maybe_warn_non_public_image_url(url: str) -> None:
    """Best-effort hint when a URL is unlikely to be fetchable by Google's
    image-retrieval servers (which never send cookies or auth headers).

    Slides `replaceImage` succeeds at the API layer even when the fetch
    fails — the slide then shows a broken-image glyph in the editor. This
    log line is the only server-side signal short of opening the deck.
    """
    if not url:
        return
    u = url.lower()
    risky = (
        "/api/file/" in u,
        "auth=" in u,
        "token=" in u,
        "localhost" in u,
        re.match(r"https?://10\.", u) is not None,
        re.match(r"https?://192\.168\.", u) is not None,
    )
    if any(risky):
        logger.warning(
            f"[_builders] Image URL may not be fetchable by Google Slides "
            f"(no cookies / auth headers): {url[:120]}{'…' if len(url) > 120 else ''}. "
            f"Use a **public HTTPS URL** (PNG/JPEG/GIF) or a Drive file shared "
            f"as \"Anyone with the link can view\" with a direct download link."
        )


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


# Compact table defaults — Google createTable *honors* size, so a near-full-slide
# height stretches every row. Prefer content-fit height + thin borders + bold
# headers (no heavy header fill) so tables match clean audit-deck typography.
_TABLE_DEFAULT_ROW_H_PT = 28.0
_TABLE_DEFAULT_FONT_SIZE_PT = 11.0
_TABLE_DEFAULT_BORDER_PT = 0.75
_TABLE_DEFAULT_BORDER_COLOR = "LIGHT2"  # themeColor — soft grid; borders can't use alpha
_TABLE_MIN_COL_W_PT = 32.0  # Slides API floor for columnWidth
_TABLE_DEFAULT_PAD_X_PT = 40.0
_TABLE_DEFAULT_PAD_Y_PT = 90.0

_THEME_COLOR_NAMES = frozenset(
    {
        "DARK1",
        "LIGHT1",
        "DARK2",
        "LIGHT2",
        "ACCENT1",
        "ACCENT2",
        "ACCENT3",
        "ACCENT4",
        "ACCENT5",
        "ACCENT6",
        "HYPERLINK",
        "FOLLOWED_HYPERLINK",
        "TEXT1",
        "TEXT2",
        "BACKGROUND1",
        "BACKGROUND2",
    }
)


def _hex_to_slides_rgb(hex_color: str) -> Dict[str, float]:
    """'#DADCE0' -> {'red': …, 'green': …, 'blue': …} for Slides solidFill/opaqueColor."""
    h = str(hex_color).lstrip("#").strip()
    if len(h) != 6:
        raise ValueError(f"Invalid HEX color '{hex_color}'. Expected #RRGGBB.")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}


def _slides_color_value(color: str) -> Dict[str, Any]:
    """Accept `#RRGGBB` or a Slides ThemeColorType name (`DARK1`, `ACCENT1`, …)."""
    raw = str(color).strip()
    if raw.startswith("#"):
        return {"rgbColor": _hex_to_slides_rgb(raw)}
    name = raw.upper().replace(" ", "_")
    if name in _THEME_COLOR_NAMES:
        return {"themeColor": name}
    raise ValueError(
        f"Invalid color '{color}'. Use #RRGGBB or a theme token "
        f"({', '.join(sorted(_THEME_COLOR_NAMES))})."
    )


def _slides_opaque_color(color: str) -> Dict[str, Any]:
    return {"opaqueColor": _slides_color_value(color)}


def _slides_solid_fill(color: str, alpha: Optional[float] = None) -> Dict[str, Any]:
    solid: Dict[str, Any] = {"color": _slides_color_value(color)}
    if alpha is not None:
        solid["alpha"] = float(alpha)
    return {"solidFill": solid}


def _normalize_table_text_style(
    style: Optional[Dict[str, Any]],
    *,
    font_family: Optional[str] = None,
    font_weight: Optional[int] = None,
    font_size_pt: Optional[float] = None,
) -> Dict[str, Any]:
    """Build a Slides TextStyle dict, accepting either API keys or friendly aliases.

    Friendly keys: font_family, font_weight (100–900), font_size, color (#RRGGBB),
    bold, italic. API keys (fontFamily, weightedFontFamily, fontSize, …) pass through.
    Stale faces like Inter are coerced to Roboto Light.
    """
    out: Dict[str, Any] = {}
    src = dict(style or {})

    def _coerce_face(name: Optional[str]) -> Optional[str]:
        if not name:
            return None
        low = str(name).strip().lower()
        if low.startswith("inter") or low == "roboto":
            return "Roboto Light"
        return str(name).strip()

    # Friendly → API
    if "font_family" in src and "fontFamily" not in src and "weightedFontFamily" not in src:
        src["fontFamily"] = src.pop("font_family")
    else:
        src.pop("font_family", None)
    if "font_weight" in src and "weightedFontFamily" not in src:
        # Defer: applied below once font family is known.
        pass
    if "font_size" in src and "fontSize" not in src:
        src["fontSize"] = {"magnitude": float(src.pop("font_size")), "unit": "PT"}
    else:
        src.pop("font_size", None)
    if "color" in src and "foregroundColor" not in src:
        try:
            src["foregroundColor"] = _slides_opaque_color(str(src.pop("color")))
        except ValueError:
            src.pop("color", None)
    else:
        src.pop("color", None)

    for key, value in src.items():
        if key in ("font_weight",) or value is None:
            continue
        out[key] = value

    # Prefer explicit brand font_family arg; coerce Inter / bare Roboto.
    family = (
        _coerce_face(font_family)
        or _coerce_face(
            (out.get("weightedFontFamily") or {}).get("fontFamily")
            or out.get("fontFamily")
        )
        or "Roboto Light"
    )

    weight = src.get("font_weight", font_weight)
    if weight is None:
        weight = 300
    try:
        weight = int(weight)
    except (TypeError, ValueError):
        weight = 300

    # Set BOTH fontFamily and weightedFontFamily (matching) — Slides UI + API
    # are most reliable when they agree.
    out["fontFamily"] = str(family)
    out["weightedFontFamily"] = {
        "fontFamily": str(family),
        "weight": weight,
    }

    if font_size_pt is not None and "fontSize" not in out:
        out["fontSize"] = {"magnitude": float(font_size_pt), "unit": "PT"}

    return out


def _estimate_col_widths(
    all_rows: List[List[Any]],
    n_cols: int,
    total_w_pt: float,
    explicit: Optional[List[float]] = None,
) -> List[float]:
    """Distribute `total_w_pt` across columns.

    Prefer explicit widths (pts or fractions summing ~1). Otherwise weight by
    max cell character length so narrative columns get more room (reduces wrap
    overflow). Every column respects the Slides 32pt minimum.
    """
    usable = max(total_w_pt, _TABLE_MIN_COL_W_PT * n_cols)

    if explicit and len(explicit) == n_cols:
        vals = [float(v) for v in explicit]
        total = sum(vals) or 1.0
        # Fractions (sum ≈ 1) vs absolute points.
        if 0.5 <= total <= 1.5:
            widths = [max(_TABLE_MIN_COL_W_PT, v / total * usable) for v in vals]
        else:
            widths = [max(_TABLE_MIN_COL_W_PT, v) for v in vals]
        # Renormalize if mins pushed us over/under.
        scale = usable / (sum(widths) or 1.0)
        return [w * scale for w in widths]

    weights = [1.0] * n_cols
    for row in all_rows:
        for c in range(n_cols):
            raw = row[c] if c < len(row) else ""
            if isinstance(raw, dict):
                text = raw.get("text")
                if text is None:
                    text = raw.get("label") or raw.get("value") or ""
            else:
                text = "" if raw is None else str(raw)
            # Soft cap so one huge cell doesn't starve siblings.
            weights[c] = max(weights[c], min(len(str(text)), 80) + 4.0)

    total_w = sum(weights) or float(n_cols)
    widths = [max(_TABLE_MIN_COL_W_PT, (w / total_w) * usable) for w in weights]
    scale = usable / (sum(widths) or 1.0)
    return [w * scale for w in widths]


def _cell_parts(value: Any) -> Tuple[str, Optional[str], Optional[Dict[str, Any]], bool]:
    """Return (text, link_url, per-cell style, is_section_row_marker).

    Section rows are dicts like `{"section": "Engagement"}` (or
    `{"row_type": "section", "text": "..."}`) and span the full table width.
    """
    if isinstance(value, dict):
        if value.get("section") is not None or value.get("row_type") == "section":
            text = value.get("section")
            if text is None:
                text = value.get("text") or value.get("label") or ""
            return (str(text), None, {"bold": True}, True)
        text = value.get("text")
        if text is None:
            text = value.get("label") or value.get("value") or ""
        link = value.get("link") or value.get("url") or value.get("href")
        cell_style = value.get("style")
        if cell_style is None and value.get("color"):
            cell_style = {"color": value["color"]}
        return (
            "" if text is None else str(text),
            str(link) if link else None,
            dict(cell_style) if isinstance(cell_style, dict) else None,
            False,
        )
    return ("" if value is None else str(value), None, None, False)


def _is_section_row(row: Any) -> bool:
    if isinstance(row, dict) and (
        row.get("section") is not None or row.get("row_type") == "section"
    ):
        return True
    if isinstance(row, list) and len(row) == 1 and isinstance(row[0], dict):
        return _is_section_row(row[0])
    return False


def _section_label(row: Any) -> str:
    if isinstance(row, dict):
        return str(row.get("section") or row.get("text") or row.get("label") or "")
    if isinstance(row, list) and row:
        return _section_label(row[0])
    return ""


_COLUMN_ROLE_WEIGHTS = {
    "label": 1.3,
    "metric": 0.85,
    "value": 0.85,
    "narrative": 2.6,
    "lecture": 2.6,
    "text": 2.0,
}
_COLUMN_ROLE_ALIGN = {
    "label": "START",
    "metric": "CENTER",
    "value": "CENTER",
    "narrative": "START",
    "lecture": "START",
    "text": "START",
}


def _widths_from_column_roles(
    roles: List[str], n_cols: int, total_w_pt: float
) -> List[float]:
    weights = []
    for i in range(n_cols):
        role = str(roles[i] if i < len(roles) else "text").strip().lower()
        weights.append(_COLUMN_ROLE_WEIGHTS.get(role, 1.0))
    total = sum(weights) or float(n_cols)
    usable = max(total_w_pt, _TABLE_MIN_COL_W_PT * n_cols)
    widths = [max(_TABLE_MIN_COL_W_PT, (w / total) * usable) for w in weights]
    scale = usable / (sum(widths) or 1.0)
    return [w * scale for w in widths]


def build_table_requests(
    slide_id: str,
    table_spec: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """Create a compact, theme-aware table on a slide and fill its cells.

    `table_spec` shape:
      {
        "headers": ["Metric", "Value"],
        "rows": [
          ["...", "..."],
          {"section": "Engagement"},          # full-width merged section row
          ["Likes", {"text": "0", "color": "#C5221F"}, "… **important** …"],
        ],
        "position": {"x": 40, "y": 95, "w": 640},  # omit h for content-fit
        "column_widths": [120, 100, 420],
        "column_roles": ["label", "metric", "narrative"],  # widths + alignment
        "column_align": ["START", "END", "START"],         # overrides roles
        "font_family": "Roboto Light",
        "font_weight": 300,
        "font_size": 11,
        "row_height": 28,
        "border_color": "#DADCE0",
        "border_weight": 0.75,
        "header_underline": true,             # thicker bottom border on header
        "header_background": null,
        "zebra": true,                        # or "zebra_color": "#F8F9FA"
        "header_style": {...},
        "body_style": {...},
      }
    """
    headers = table_spec.get("headers") or []
    raw_body_rows = table_spec.get("rows") or []
    if not headers and not raw_body_rows:
        return []

    # Normalize body rows: section dicts → single-cell marker lists for sizing,
    # but we track which logical body indices are sections for merge requests.
    body_rows: List[List[Any]] = []
    section_body_indices: List[int] = []
    for r in raw_body_rows:
        if _is_section_row(r):
            section_body_indices.append(len(body_rows))
            body_rows.append([{"section": _section_label(r)}])
        else:
            body_rows.append(list(r) if isinstance(r, (list, tuple)) else [r])

    if headers and body_rows:
        all_rows: List[List[Any]] = [list(headers)] + body_rows
        header_offset = 1
    elif headers:
        all_rows = [list(headers)]
        header_offset = 1
    else:
        all_rows = body_rows
        header_offset = 0

    n_rows = len(all_rows)
    n_cols = max(
        (len(headers) if headers else 0),
        max((len(r) for r in body_rows), default=1),
        1,
    )
    # Pad short rows so createTable grid is rectangular; section rows stay
    # conceptually 1-cell but we fill placeholders (merged afterward).
    normalized_rows: List[List[Any]] = []
    for r_idx, row in enumerate(all_rows):
        is_section = header_offset and r_idx >= header_offset and (
            (r_idx - header_offset) in section_body_indices
        )
        if is_section:
            label = _section_label(row[0] if row else "")
            normalized_rows.append([label] + [""] * (n_cols - 1))
        else:
            padded = list(row) + [""] * (n_cols - len(row))
            normalized_rows.append(padded[:n_cols])
    all_rows = normalized_rows

    row_h = float(table_spec.get("row_height") or _TABLE_DEFAULT_ROW_H_PT)
    auto_h = max(row_h * n_rows, row_h)
    default_pos = {
        "x": _TABLE_DEFAULT_PAD_X_PT,
        "y": _TABLE_DEFAULT_PAD_Y_PT,
        "w": DEFAULT_PAGE_W_PT - (2 * _TABLE_DEFAULT_PAD_X_PT),
        "h": auto_h,
    }
    pos = _position(table_spec.get("position"), default_pos)
    if not (table_spec.get("position") or {}).get("h"):
        pos["h"] = min(auto_h, DEFAULT_PAGE_H_PT - pos["y"] - 20.0)

    font_family = table_spec.get("font_family") or table_spec.get("fontFamily") or "Roboto Light"
    if str(font_family).strip().lower().startswith("inter") or str(font_family).strip().lower() == "roboto":
        font_family = "Roboto Light"
    font_weight = table_spec.get("font_weight")
    if font_weight is not None:
        try:
            font_weight = int(font_weight)
        except (TypeError, ValueError):
            font_weight = 300
    else:
        font_weight = 300
    font_size = table_spec.get("font_size")
    font_size = float(font_size) if font_size is not None else _TABLE_DEFAULT_FONT_SIZE_PT

    header_style = _normalize_table_text_style(
        table_spec.get("header_style") or {"bold": True},
        font_family=font_family,
        font_weight=(
            int((table_spec.get("header_style") or {}).get("font_weight"))
            if (table_spec.get("header_style") or {}).get("font_weight") is not None
            else 300
        ),
        font_size_pt=font_size,
    )
    if "bold" not in header_style:
        header_style["bold"] = True

    body_style = _normalize_table_text_style(
        table_spec.get("body_style"),
        font_family=font_family,
        font_weight=font_weight,
        font_size_pt=font_size,
    )
    section_style = _normalize_table_text_style(
        {"bold": True},
        font_family=font_family,
        font_weight=600 if not font_weight or font_weight < 600 else font_weight,
        font_size_pt=font_size,
    )

    # Theme-linked text color (DARK1 by default) — omit if caller already set it.
    text_color = table_spec.get("text_color") or "DARK1"
    for style_dict in (header_style, body_style, section_style):
        if "foregroundColor" not in style_dict:
            try:
                style_dict["foregroundColor"] = _slides_opaque_color(str(text_color))
            except ValueError:
                pass

    border_color = table_spec.get("border_color") or _TABLE_DEFAULT_BORDER_COLOR
    border_w = float(table_spec.get("border_weight") or _TABLE_DEFAULT_BORDER_PT)
    header_bg = table_spec.get("header_background")
    header_bg_alpha = table_spec.get("header_background_alpha")
    header_underline = table_spec.get("header_underline", True)
    header_underline_color = table_spec.get("header_underline_color") or "DARK1"
    zebra = table_spec.get("zebra") or table_spec.get("zebra_rows")
    zebra_color = table_spec.get("zebra_color") or "LIGHT2"
    section_bg = table_spec.get("section_background") or "LIGHT2"

    roles = table_spec.get("column_roles") or []
    if roles and not table_spec.get("column_widths"):
        col_widths = _widths_from_column_roles([str(r) for r in roles], n_cols, pos["w"])
    else:
        col_widths = _estimate_col_widths(
            all_rows, n_cols, pos["w"], explicit=table_spec.get("column_widths")
        )

    # Per-column paragraph alignment.
    col_align: List[str] = []
    explicit_align = table_spec.get("column_align") or table_spec.get("column_aligns") or []
    for c in range(n_cols):
        if c < len(explicit_align) and explicit_align[c]:
            col_align.append(str(explicit_align[c]).upper())
        elif c < len(roles):
            col_align.append(_COLUMN_ROLE_ALIGN.get(str(roles[c]).lower(), "START"))
        else:
            col_align.append("START")

    table_id = gen_id("tbl")
    # IMPORTANT: The Slides API *ignores* size/transform on createTable and
    # centers the table on the slide. We must follow up with
    # updatePageElementTransform (translate only — tables cannot be scaled)
    # plus column widths for horizontal size.
    requests: List[Dict[str, Any]] = [
        {
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": slide_id,
                },
                "rows": n_rows,
                "columns": n_cols,
            }
        },
        {
            "updatePageElementTransform": {
                "objectId": table_id,
                "applyMode": "ABSOLUTE",
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "shearX": 0,
                    "shearY": 0,
                    "translateX": float(pos["x"]),
                    "translateY": float(pos["y"]),
                    "unit": "PT",
                },
            }
        },
    ]

    requests.append(
        {
            "updateTableRowProperties": {
                "objectId": table_id,
                "tableRowProperties": {"minRowHeight": _pt(row_h)},
                "fields": "minRowHeight",
            }
        }
    )

    for c_idx, width in enumerate(col_widths):
        requests.append(
            {
                "updateTableColumnProperties": {
                    "objectId": table_id,
                    "columnIndices": [c_idx],
                    "tableColumnProperties": {
                        "columnWidth": _pt(max(_TABLE_MIN_COL_W_PT, width))
                    },
                    "fields": "columnWidth",
                }
            }
        )

    try:
        # Table borders reject alpha other than 0 or 1 — always solid.
        border_fill = _slides_solid_fill(str(border_color))
    except ValueError:
        border_fill = _slides_solid_fill(_TABLE_DEFAULT_BORDER_COLOR)
    requests.append(
        {
            "updateTableBorderProperties": {
                "objectId": table_id,
                "borderPosition": "ALL",
                "tableBorderProperties": {
                    "tableBorderFill": border_fill,
                    "weight": _pt(border_w),
                    "dashStyle": "SOLID",
                },
                "fields": "tableBorderFill,weight,dashStyle",
            }
        }
    )

    # Editorial header rule: thicker bottom border under the header row.
    if headers and header_underline:
        try:
            underline_fill = _slides_solid_fill(str(header_underline_color))
        except ValueError:
            underline_fill = border_fill
        requests.append(
            {
                "updateTableBorderProperties": {
                    "objectId": table_id,
                    "tableRange": {
                        "location": {"rowIndex": 0, "columnIndex": 0},
                        "rowSpan": 1,
                        "columnSpan": n_cols,
                    },
                    "borderPosition": "BOTTOM",
                    "tableBorderProperties": {
                        "tableBorderFill": underline_fill,
                        "weight": _pt(max(border_w * 2, 1.25)),
                        "dashStyle": "SOLID",
                    },
                    "fields": "tableBorderFill,weight,dashStyle",
                }
            }
        )

    requests.append(
        {
            "updateTableCellProperties": {
                "objectId": table_id,
                "tableCellProperties": {"contentAlignment": "TOP"},
                "fields": "contentAlignment",
            }
        }
    )
    if headers and header_bg:
        try:
            alpha = None
            # Solid hex fills must stay opaque — ACCENT1@0.14 was reading as beige.
            if header_bg_alpha is not None and not str(header_bg).startswith("#"):
                alpha = float(header_bg_alpha)
            fill = _slides_solid_fill(str(header_bg), alpha=alpha)
            fields = (
                "tableCellBackgroundFill.solidFill"
                if alpha is not None
                else "tableCellBackgroundFill.solidFill.color"
            )
            requests.append(
                {
                    "updateTableCellProperties": {
                        "objectId": table_id,
                        "tableRange": {
                            "location": {"rowIndex": 0, "columnIndex": 0},
                            "rowSpan": 1,
                            "columnSpan": n_cols,
                        },
                        "tableCellProperties": {
                            "tableCellBackgroundFill": fill,
                        },
                        "fields": fields,
                    }
                }
            )
        except (ValueError, TypeError):
            pass

    # Zebra striping on body rows (skip section rows).
    if zebra:
        try:
            zebra_fill = _slides_solid_fill(str(zebra_color))
            for body_i in range(len(body_rows)):
                if body_i in section_body_indices:
                    continue
                if body_i % 2 == 1:
                    r_idx = body_i + header_offset
                    requests.append(
                        {
                            "updateTableCellProperties": {
                                "objectId": table_id,
                                "tableRange": {
                                    "location": {"rowIndex": r_idx, "columnIndex": 0},
                                    "rowSpan": 1,
                                    "columnSpan": n_cols,
                                },
                                "tableCellProperties": {
                                    "tableCellBackgroundFill": zebra_fill,
                                },
                                "fields": "tableCellBackgroundFill.solidFill.color",
                            }
                        }
                    )
        except ValueError:
            pass

    # Soft fill for section rows.
    for body_i in section_body_indices:
        r_idx = body_i + header_offset
        try:
            requests.append(
                {
                    "updateTableCellProperties": {
                        "objectId": table_id,
                        "tableRange": {
                            "location": {"rowIndex": r_idx, "columnIndex": 0},
                            "rowSpan": 1,
                            "columnSpan": n_cols,
                        },
                        "tableCellProperties": {
                            "tableCellBackgroundFill": _slides_solid_fill(
                                str(section_bg), alpha=0.55
                            ),
                        },
                        "fields": "tableCellBackgroundFill.solidFill.color",
                    }
                }
            )
        except ValueError:
            pass
        if n_cols > 1:
            requests.append(
                {
                    "mergeTableCells": {
                        "objectId": table_id,
                        "tableRange": {
                            "location": {"rowIndex": r_idx, "columnIndex": 0},
                            "rowSpan": 1,
                            "columnSpan": n_cols,
                        },
                    }
                }
            )

    for r_idx, row in enumerate(all_rows):
        is_header = bool(headers and r_idx == 0)
        is_section = (
            header_offset
            and r_idx >= header_offset
            and (r_idx - header_offset) in section_body_indices
        )
        # Section rows: only fill the first (merged) cell.
        for c_idx in (range(1) if is_section else range(n_cols)):
            value = row[c_idx] if c_idx < len(row) else ""
            text, link, per_cell, _ = _cell_parts(value)
            if not text:
                continue
            plain, bold_ranges = _parse_inline_bold(text)
            requests.append(
                {
                    "insertText": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
                        "text": plain,
                        "insertionIndex": 0,
                    }
                }
            )
            if is_section:
                base = section_style
            elif is_header:
                base = header_style
            else:
                base = body_style
            cell_style = dict(base)
            if per_cell:
                cell_style.update(
                    _normalize_table_text_style(
                        per_cell,
                        font_family=font_family,
                        font_size_pt=font_size,
                    )
                )
            # Keep fontFamily + weightedFontFamily in sync (API requires match).
            wff = cell_style.get("weightedFontFamily") or {}
            if wff.get("fontFamily"):
                cell_style["fontFamily"] = wff["fontFamily"]
            if link:
                cell_style["link"] = {"url": link}
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
            for start, end in bold_ranges:
                if end <= start:
                    continue
                requests.append(
                    {
                        "updateTextStyle": {
                            "objectId": table_id,
                            "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
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
            alignment = "START" if (is_header or is_section) else col_align[c_idx]
            requests.append(
                {
                    "updateParagraphStyle": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
                        "textRange": {"type": "ALL"},
                        "style": {"alignment": alignment},
                        "fields": "alignment",
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
    _maybe_warn_non_public_image_url(url)
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
    linking_mode: str = "NOT_LINKED_IMAGE",
) -> List[Dict[str, Any]]:
    """Embed a Google Sheets chart on a slide via its chartId.

    `linking_mode` defaults to ``NOT_LINKED_IMAGE`` (static snapshot at embed
    time). This avoids the common failure mode where ``LINKED`` charts render
    as a broken placeholder for viewers who cannot access the source
    spreadsheet, or when Workspace link policies block live chart resolution.
    Pass ``LINKED`` only when you need live refresh and every deck viewer has
    read access to the data sheet.
    """
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
    speaker_notes_object_id: str,
    notes_text: str,
    existing_notes_text: str = "",
) -> List[Dict[str, Any]]:
    """Set speaker notes on the notes shape, replacing any existing text.

    Crucially: only emits a `deleteText` request if `existing_notes_text` is
    non-empty. The Slides API rejects `deleteText` with
    `textRange: {type: "ALL"}` on an empty shape because it internally
    resolves `ALL` to `(startIndex: 0, endIndex: 0)` and then enforces
    `endIndex > startIndex` ("HTTP 400: The startIndex 0 must be less
    than the endIndex 0"). For freshly-created slides the notes shape is
    always empty, so the delete must be skipped.

    Pass `existing_notes_text=""` (the default) when you know the notes
    shape was just created. Pass the live text otherwise — the caller can
    extract it from the slide's `notesPage.pageElements[].shape.text` in
    the same `presentations.get` it used to find `speakerNotesObjectId`.
    """
    if not notes_text:
        return []
    requests: List[Dict[str, Any]] = []
    if existing_notes_text:
        requests.append(
            {
                "deleteText": {
                    "objectId": speaker_notes_object_id,
                    "textRange": {"type": "ALL"},
                }
            }
        )
    requests.append(
        {
            "insertText": {
                "objectId": speaker_notes_object_id,
                "insertionIndex": 0,
                "text": notes_text,
            }
        }
    )
    return requests


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

    # Discover the layout's REAL placeholder indexes. This is critical
    # because the Slides backend returns either HTTP 400 ('placeholder is
    # not on the page') or HTTP 500 ('Internal error encountered') if a
    # `placeholderIdMapping` references a type/index combo the layout
    # doesn't actually have. We do this for BOTH custom layouts AND
    # predefined ones (predefined layout names like `TITLE_AND_BODY` keep
    # their name in customized templates, but a user might have deleted /
    # reorganized their placeholders, which breaks the naive
    # 'predefined layouts always have index 0' assumption).
    layout_placeholders_by_type: Dict[str, List[int]] = {}
    discovered_layout_id: Optional[str] = None
    if "layoutId" in layout_ref:
        discovered_layout_id = layout_ref["layoutId"]
    elif "predefinedLayout" in layout_ref:
        discovered_layout_id = find_predefined_layout_id(
            presentation, layout_ref["predefinedLayout"]
        )
    if discovered_layout_id:
        layout_placeholders_by_type = get_layout_placeholders_by_type(
            presentation, discovered_layout_id
        )

    def _is_multi_occurrence(ph_type: str) -> bool:
        """The resolved layout has 2+ placeholders of `ph_type`?

        Triggers the deferred-rebind workaround for the Slides multi-BODY
        ghost-bug. Applies to both custom AND predefined layouts that have
        multiple same-type placeholders (e.g. predefined
        `TITLE_AND_TWO_COLUMNS` has 2x BODY).
        """
        if not layout_placeholders_by_type:
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

        If we couldn't discover the layout's actual placeholders (predefined
        layout absent from `presentation.layouts[]`), fall back to the legacy
        type+index behavior — better to attempt the binding than to skip the
        slide entirely.
        """
        if not layout_placeholders_by_type:
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
        if layout_placeholders_by_type and occurrence >= len(indexes):
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
            discovered_layout_id
            and i >= 1
            and len(layout_placeholders_by_type.get("BODY") or []) > 1
        ):
            geom = get_layout_placeholder_geometry(
                presentation, discovered_layout_id, "BODY", i
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

    # Title + Table (and similar): layout still exposes a BODY placeholder.
    # If the deck only supplies a `table` (no body text), that BODY stays as
    # the grey "Click to add text" prompt behind the table. Bind + delete it.
    has_table = bool(slide_spec.get("table"))
    if has_table and not body_texts:
        n_body = len(layout_placeholders_by_type.get("BODY") or [])
        # If we couldn't discover placeholders, still try BODY occurrence 0.
        to_clear = n_body if n_body else 1
        for i in range(to_clear):
            allocated = _allocate_placeholder("BODY", i, f"body_unused[{i}]")
            if allocated is None and n_body == 0 and i == 0:
                # Predefined / unknown layout: attempt a type-only mapping.
                ph_id = gen_id("ph")
                placeholder_ids[f"body_unused[{i}]"] = ph_id
                placeholder_mappings.append(
                    {
                        "layoutPlaceholder": {"type": "BODY"},
                        "objectId": ph_id,
                    }
                )

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

    # Remove unused BODY placeholders so "Click to add text" never shows
    # behind tables / charts that occupy the body area.
    for key, ph_id in list(placeholder_ids.items()):
        if key.startswith("body_unused"):
            content_requests.append({"deleteObject": {"objectId": ph_id}})

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
        url = spec["url"]
        _maybe_warn_non_public_image_url(url)
        method = _normalize_image_replace_method(spec.get("method"))
        content_requests.append(
            {
                "replaceImage": {
                    "imageObjectId": ph_id,
                    "url": url,
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
        table_spec = dict(slide_spec["table"])
        # Fit the table into the layout's BODY area. Agent-supplied `position`
        # often parks the table inside the dark brand bar — always snap to BODY
        # geometry when available (Title + Table and similar).
        if discovered_layout_id:
            geom = get_layout_placeholder_geometry(
                presentation, discovered_layout_id, "BODY", 0
            )
            if geom is not None:
                size, transform = geom
                box = geometry_to_position(size, transform)
                # BODY on Title + Table already sits below the dark brand bar.
                # Tiny pad only — avoid a dead white gap above the table.
                side_inset = 12.0
                top_inset = 8.0
                if box["w"] > 2 * side_inset and box["h"] > top_inset + 20:
                    table_spec["position"] = {
                        "x": box["x"] + side_inset,
                        "y": box["y"] + top_inset,
                        "w": box["w"] - 2 * side_inset,
                    }
        content_requests.extend(build_table_requests(slide_id, table_spec))

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
