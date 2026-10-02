"""Deck preprocessing for `create_audit_presentation`.

Pure helpers (no I/O except optional image HEAD checks):
  * capacity findings collection (returned to the MCP client)
  * theme font extraction from template masters
  * text_defaults → placeholder styles + table/chart font inheritance
  * chart position inference from layout name
  * auto-split of oversized tables and long bodies
  * image URL preflight (drop clearly-unreachable URLs before build)
"""
from __future__ import annotations

import asyncio
import copy
import logging
import re
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Mirrors audit_builder._CAPACITY_LIMITS — keep soft targets here for auto-split.
_TABLE_ROWS_SOFT = 8
_BODY_SOFT = {
    "Title + Body": 500,
    "TITLE_AND_BODY": 500,
    "Title + Chart + Body": 300,
    "Two Columns": 250,
    "Title + Two Columns": 250,
}
_BODY_SOFT_DEFAULT = 500

_CHART_POSITION_BY_LAYOUT = {
    "title + chart + body": {"x": 380.0, "y": 110.0, "w": 300.0, "h": 250.0},
    "title + chart": {"x": 60.0, "y": 110.0, "w": 600.0, "h": 270.0},
}

@dataclass
class PreprocessResult:
    deck: Dict[str, Any]
    capacity_findings: List[Dict[str, Any]] = field(default_factory=list)
    split_notes: List[str] = field(default_factory=list)
    image_preflight: List[Dict[str, Any]] = field(default_factory=list)
    theme_font: Optional[str] = None
    inferred_chart_positions: int = 0


def _normalize_layout(name: str) -> str:
    return " ".join(str(name or "").strip().lower().split())


@dataclass
class ThemeInfo:
    """Snapshot of the template master typography + color scheme."""

    font_family: Optional[str] = None
    font_weight: Optional[int] = None
    body_font_size: Optional[float] = None
    title_font_size: Optional[float] = None
    # ThemeColorType → rgb floats (for Sheets charts that can't use themeColor refs)
    colors: Dict[str, Dict[str, float]] = field(default_factory=dict)

    def as_dict(self) -> Dict[str, Any]:
        return {
            "font_family": self.font_family,
            "font_weight": self.font_weight,
            "body_font_size": self.body_font_size,
            "title_font_size": self.title_font_size,
            "colors": self.colors,
        }


def _placeholder_type(element: Dict[str, Any]) -> Optional[str]:
    shape = element.get("shape") or {}
    ph = shape.get("placeholder") or {}
    if ph.get("type"):
        return str(ph["type"])
    image = element.get("image") or {}
    ph = image.get("placeholder") or {}
    return str(ph["type"]) if ph.get("type") else None


def _first_text_style(element: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    shape = element.get("shape") or {}
    text = shape.get("text") or {}
    for te in text.get("textElements") or []:
        style = te.get("textStyle")
        if style:
            return style
    return None


def extract_theme(presentation: Dict[str, Any]) -> ThemeInfo:
    """Read BODY/TITLE fonts + colorScheme from the template master/layouts.

    Tables and free TEXT_BOX overlays cannot inherit live theme styles (Slides
    API limitation), so we snapshot these values at build time. Placeholder
    text on custom layouts still inherits the master unless the deck JSON
    overrides `styles`.
    """
    theme = ThemeInfo()

    def _ingest_style(style: Dict[str, Any], *, prefer_body: bool) -> None:
        wff = style.get("weightedFontFamily") or {}
        family = wff.get("fontFamily") or style.get("fontFamily")
        weight = wff.get("weight")
        size = (style.get("fontSize") or {}).get("magnitude")
        if family and (prefer_body or not theme.font_family):
            theme.font_family = str(family)
            if weight is not None:
                try:
                    theme.font_weight = int(weight)
                except (TypeError, ValueError):
                    pass
        if size is not None:
            try:
                size_f = float(size)
            except (TypeError, ValueError):
                return
            if prefer_body and theme.body_font_size is None:
                theme.body_font_size = size_f
            elif not prefer_body and theme.title_font_size is None:
                theme.title_font_size = size_f

    def _scan(elements: List[Dict[str, Any]]) -> None:
        # Prefer BODY placeholders — that's the table/body brand face.
        bodies: List[Dict[str, Any]] = []
        titles: List[Dict[str, Any]] = []
        other: List[Dict[str, Any]] = []
        for el in elements or []:
            style = _first_text_style(el)
            if not style:
                continue
            ph = _placeholder_type(el)
            if ph == "BODY":
                bodies.append(style)
            elif ph in ("TITLE", "CENTERED_TITLE", "SUBTITLE"):
                titles.append(style)
            else:
                other.append(style)
        for style in bodies:
            _ingest_style(style, prefer_body=True)
        for style in titles:
            _ingest_style(style, prefer_body=False)
        if not theme.font_family:
            for style in other:
                _ingest_style(style, prefer_body=True)

    for master in presentation.get("masters") or []:
        props = master.get("pageProperties") or {}
        scheme = props.get("colorScheme") or {}
        for entry in scheme.get("colors") or []:
            ctype = entry.get("type")
            rgb = (entry.get("color") or {}).get("rgbColor")
            if ctype and isinstance(rgb, dict):
                theme.colors[str(ctype)] = {
                    "red": float(rgb.get("red") or 0),
                    "green": float(rgb.get("green") or 0),
                    "blue": float(rgb.get("blue") or 0),
                }
        _scan(master.get("pageElements") or [])
    if not theme.font_family:
        for layout in presentation.get("layouts") or []:
            _scan(layout.get("pageElements") or [])
            if theme.font_family:
                break
    # Masters/layouts often omit explicit fontFamily (inherited from the
    # Slides theme UI). Fall back to any styled text on existing slides.
    if not theme.font_family:
        for slide in presentation.get("slides") or []:
            _scan(slide.get("pageElements") or [])
            if theme.font_family:
                break
    return theme


def extract_theme_font(presentation: Dict[str, Any]) -> Optional[str]:
    """Backward-compatible helper — prefer `extract_theme`."""
    return extract_theme(presentation).font_family


def _theme_series_colors(theme: ThemeInfo) -> List[str]:
    """Resolve ACCENT1–6 from the master colorScheme to #RRGGBB for Sheets charts."""
    out: List[str] = []
    for key in ("ACCENT1", "ACCENT2", "ACCENT3", "ACCENT4", "ACCENT5", "ACCENT6"):
        rgb = theme.colors.get(key)
        if not rgb:
            continue
        r = int(round(rgb["red"] * 255))
        g = int(round(rgb["green"] * 255))
        b = int(round(rgb["blue"] * 255))
        out.append(f"#{r:02X}{g:02X}{b:02X}")
    return out


def _text_style_from_defaults(
    text_defaults: Dict[str, Any],
    *,
    role: str = "body",
) -> Dict[str, Any]:
    """Build a Slides TextStyle dict from explicit deck.text_defaults."""
    style: Dict[str, Any] = {}
    family = text_defaults.get("font_family") or text_defaults.get("fontFamily")
    weight = text_defaults.get("font_weight")
    # Slides has no "Roboto Light" family — map to Roboto + weight 300.
    if family and str(family).strip().lower() in ("roboto light", "robotolight"):
        family = "Roboto"
        if weight is None:
            weight = 300
    size_key = "title_font_size" if role == "title" else "body_font_size"
    size = text_defaults.get(size_key)
    if size is None and role == "body":
        size = text_defaults.get("font_size")
    if family and weight is not None:
        try:
            style["weightedFontFamily"] = {
                "fontFamily": str(family),
                "weight": int(weight),
            }
        except (TypeError, ValueError):
            style["fontFamily"] = str(family)
    elif family:
        style["fontFamily"] = str(family)
    if size is not None:
        try:
            style["fontSize"] = {"magnitude": float(size), "unit": "PT"}
        except (TypeError, ValueError):
            pass
    return style


def apply_text_defaults(
    deck: Dict[str, Any],
    theme: Optional[ThemeInfo] = None,
    theme_font: Optional[str] = None,  # legacy kwarg
) -> Dict[str, Any]:
    """Apply typography with **theme-first** semantics.

    * Placeholders (TITLE/BODY on the layout) inherit the master **unless** the
      deck explicitly sets `text_defaults` or per-slide `styles`. We deliberately
      do NOT stamp theme fonts onto every placeholder — that would freeze the
      font at build time and ignore later theme edits in the Slides editor.
    * Tables / chart embeds / Two-Columns body[1] overlays **cannot** inherit
      live theme styles (API limitation). For those we snapshot the template
      font/colors into `table_defaults` / overlay styles at build time.
    """
    out = copy.deepcopy(deck)
    if theme is None and theme_font:
        theme = ThemeInfo(font_family=theme_font)
    theme = theme or ThemeInfo()

    user_text_defaults = out.get("text_defaults")
    explicit_text = isinstance(user_text_defaults, dict) and bool(user_text_defaults)
    text_defaults = dict(user_text_defaults or {})

    # Agent-owned typography when set. Tables + TEXT_BOX overlays cannot inherit
    # the live theme — soft-default Roboto Light when the agent omitted fonts
    # (otherwise Slides falls back to Arial). No Inter→Roboto coercion.
    brand_family = (
        text_defaults.get("font_family")
        or (out.get("table_defaults") or {}).get("font_family")
        or (out.get("chart_defaults") or {}).get("font_family")
    )
    brand_weight = text_defaults.get("font_weight")
    if brand_weight is None:
        brand_weight = (out.get("table_defaults") or {}).get("font_weight")
    if brand_family and str(brand_family).strip().lower() in ("roboto light", "robotolight"):
        brand_family = "Roboto"
        if brand_weight is None:
            brand_weight = 300
    if not brand_family:
        brand_family = "Roboto"
        if brand_weight is None:
            brand_weight = 300
    body_size = (
        text_defaults.get("body_font_size")
        or text_defaults.get("font_size")
        or theme.body_font_size
        or 11
    )

    # --- table_defaults: seed fonts (soft) + theme colors ---
    table_defaults = dict(out.get("table_defaults") or {})
    if not table_defaults.get("font_family"):
        table_defaults["font_family"] = brand_family
    if "font_weight" not in table_defaults:
        table_defaults["font_weight"] = (
            brand_weight if brand_weight is not None else 300
        )
    if body_size is not None and "font_size" not in table_defaults:
        try:
            table_defaults["font_size"] = min(float(body_size), 12.0)
        except (TypeError, ValueError):
            pass
    # Theme-linked colors (Slides themeColor tokens — track theme edits on rebuild).
    if not table_defaults.get("border_color"):
        table_defaults["border_color"] = "LIGHT2"
    if not table_defaults.get("zebra_color") and table_defaults.get("zebra"):
        table_defaults["zebra_color"] = "LIGHT2"
    if not table_defaults.get("text_color"):
        table_defaults["text_color"] = "DARK1"
    if not table_defaults.get("header_underline_color"):
        table_defaults["header_underline_color"] = "DARK1"
    if not table_defaults.get("section_background"):
        table_defaults["section_background"] = "LIGHT2"
    # Soft mint header (goal look). Avoid ACCENT1@low-alpha — reads as beige
    # when ACCENT1 is a dark brand green.
    user_header_bg = "header_background" in (out.get("table_defaults") or {})
    user_header_alpha = "header_background_alpha" in (out.get("table_defaults") or {})
    if not user_header_bg:
        table_defaults["header_background"] = "#E8F5E9"
        table_defaults.pop("header_background_alpha", None)
    elif (
        not user_header_alpha
        and isinstance(table_defaults.get("header_background"), str)
        and str(table_defaults["header_background"]).startswith("#")
    ):
        table_defaults.pop("header_background_alpha", None)
    out["table_defaults"] = table_defaults

    # --- chart_defaults: font + accent series when omitted ---
    chart_defaults = dict(out.get("chart_defaults") or {})
    if brand_family and not chart_defaults.get("font_family"):
        chart_defaults["font_family"] = brand_family
    if not chart_defaults.get("series_colors"):
        accents = _theme_series_colors(theme)
        if accents:
            chart_defaults["series_colors"] = accents
    if chart_defaults:
        out["chart_defaults"] = chart_defaults

    # --- Placeholders: only override when the deck explicitly asked ---
    body_style = _text_style_from_defaults(text_defaults, role="body") if explicit_text else {}
    title_style = _text_style_from_defaults(text_defaults, role="title") if explicit_text else {}

    # Overlay / column style: Roboto Light for TEXT_BOX + Two Columns bodies
    # (placeholders may inherit Regular; overlays never inherit the theme).
    light_style: Dict[str, Any] = {}
    if brand_family:
        if brand_weight is not None:
            light_style["weightedFontFamily"] = {
                "fontFamily": str(brand_family),
                "weight": int(brand_weight),
            }
            light_style["fontFamily"] = str(brand_family)
        else:
            light_style["fontFamily"] = str(brand_family)
        if body_size is not None:
            light_style["fontSize"] = {"magnitude": float(body_size), "unit": "PT"}
    overlay_style = dict(light_style)

    for slide in out.get("slides") or []:
        styles = dict(slide.get("styles") or {})
        fields_body = (slide.get("fields") or {}).get("body")

        if explicit_text:
            if title_style and "title" not in styles:
                styles["title"] = dict(title_style)
            if body_style:
                existing_body = styles.get("body")
                if existing_body is None:
                    if isinstance(fields_body, list):
                        styles["body"] = [dict(body_style) for _ in fields_body]
                    else:
                        styles["body"] = dict(body_style)
                elif isinstance(existing_body, list):
                    merged = []
                    for item in existing_body:
                        if item is None:
                            merged.append(dict(body_style))
                        elif isinstance(item, dict):
                            base = dict(body_style)
                            base.update(item)
                            merged.append(base)
                        else:
                            merged.append(item)
                    styles["body"] = merged
                elif isinstance(existing_body, dict):
                    base = dict(body_style)
                    base.update(existing_body)
                    styles["body"] = base
        elif isinstance(fields_body, list) and len(fields_body) > 1 and light_style:
            # Two Columns: stamp Roboto Light on BOTH columns. body[0] can
            # inherit a Regular face from the master; body[1] is a TEXT_BOX
            # that never inherits. Match table_defaults weight 300 on both.
            existing = styles.get("body")
            n = len(fields_body)

            def _ensure_light(item: Any) -> Dict[str, Any]:
                if not isinstance(item, dict):
                    return dict(light_style)
                if item.get("fontFamily") or item.get("weightedFontFamily"):
                    # Keep face if set, but force weight 300 when missing.
                    merged = dict(light_style)
                    merged.update(item)
                    wff = merged.get("weightedFontFamily") or {}
                    family = wff.get("fontFamily") or merged.get("fontFamily") or brand_family
                    weight = wff.get("weight")
                    if weight is None:
                        weight = brand_weight if brand_weight is not None else 300
                    merged["fontFamily"] = str(family)
                    merged["weightedFontFamily"] = {
                        "fontFamily": str(family),
                        "weight": int(weight),
                    }
                    return merged
                base = dict(light_style)
                base.update(item)
                return base

            if existing is None:
                styles["body"] = [_ensure_light(None) for _ in range(n)]
            elif isinstance(existing, list):
                merged = list(existing) + [None] * (n - len(existing))
                styles["body"] = [_ensure_light(merged[i]) for i in range(n)]
            elif isinstance(existing, dict):
                styles["body"] = [_ensure_light(existing) for _ in range(n)]

        # Do NOT stamp TITLE fonts — placeholders inherit the template theme.
        # Agent can set styles.title / text_defaults when a one-off override is needed.

        if styles:
            slide["styles"] = styles
    return out


def infer_chart_positions(deck: Dict[str, Any]) -> Tuple[Dict[str, Any], int]:
    """Fill chart.position from layout name when the agent omitted it."""
    out = copy.deepcopy(deck)
    inferred = 0
    for slide in out.get("slides") or []:
        chart = slide.get("chart")
        if not isinstance(chart, dict):
            continue
        if chart.get("position"):
            continue
        layout_key = _normalize_layout(slide.get("layout") or "")
        recipe = _CHART_POSITION_BY_LAYOUT.get(layout_key)
        if recipe:
            chart["position"] = dict(recipe)
            inferred += 1
    return out, inferred


def _suffix_title(title: Any, part: int, total: int) -> str:
    base = str(title or "").strip()
    # Strip a pre-existing (k/n) suffix so re-splits stay clean.
    base = re.sub(r"\s*\(\d+/\d+\)\s*$", "", base).strip()
    if not base:
        base = "Suite"
    return f"{base} ({part}/{total})"


def _split_body_text(text: str, soft_limit: int) -> List[str]:
    """Split a long body into chunks near soft_limit, preferring paragraph breaks."""
    text = str(text or "").strip()
    if len(text) <= soft_limit:
        return [text] if text else []

    paragraphs = re.split(r"\n\s*\n", text)
    chunks: List[str] = []
    current = ""
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        candidate = f"{current}\n\n{para}".strip() if current else para
        if len(candidate) <= soft_limit:
            current = candidate
            continue
        if current:
            chunks.append(current)
            current = ""
        if len(para) <= soft_limit:
            current = para
            continue
        # Hard-split a single oversized paragraph on sentence / line boundaries.
        lines = re.split(r"(?<=[.!?])\s+|\n", para)
        buf = ""
        for line in lines:
            line = line.strip()
            if not line:
                continue
            cand = f"{buf} {line}".strip() if buf else line
            if len(cand) <= soft_limit:
                buf = cand
            else:
                if buf:
                    chunks.append(buf)
                if len(line) <= soft_limit:
                    buf = line
                else:
                    # Absolute last resort: character slices.
                    for i in range(0, len(line), soft_limit):
                        piece = line[i : i + soft_limit].strip()
                        if piece:
                            chunks.append(piece)
                    buf = ""
        if buf:
            current = buf
    if current:
        chunks.append(current)
    return chunks or [text[:soft_limit]]


def auto_split_slides(deck: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    """Expand oversized tables / bodies into consecutive slides.

    Controlled by `deck.auto_split` (default True). Two-column bodies are not
    auto-split (structure is intentional); only string bodies are.
    """
    if deck.get("auto_split") is False:
        return copy.deepcopy(deck), []

    out = copy.deepcopy(deck)
    notes: List[str] = []
    new_slides: List[Dict[str, Any]] = []

    for slide in out.get("slides") or []:
        layout = slide.get("layout") or "BLANK"
        table = slide.get("table") if isinstance(slide.get("table"), dict) else None
        fields = slide.get("fields") or {}
        body = fields.get("body")

        # --- Table split ---
        if table:
            headers = list(table.get("headers") or [])
            rows = list(table.get("rows") or [])
            # Section rows count toward the budget like data rows.
            if len(rows) > _TABLE_ROWS_SOFT:
                chunk_size = _TABLE_ROWS_SOFT
                chunks = [rows[i : i + chunk_size] for i in range(0, len(rows), chunk_size)]
                total = len(chunks)
                title = fields.get("title") or slide.get("title") or "Tableau"
                for part_i, chunk in enumerate(chunks, start=1):
                    part = copy.deepcopy(slide)
                    part_fields = dict(part.get("fields") or {})
                    part_fields["title"] = _suffix_title(title, part_i, total)
                    part["fields"] = part_fields
                    part_table = dict(part.get("table") or {})
                    part_table["headers"] = headers
                    part_table["rows"] = chunk
                    part["table"] = part_table
                    if part_i > 1:
                        part.pop("chart", None)
                        part.pop("image", None)
                        part.pop("images", None)
                        part.pop("image_placeholders", None)
                    new_slides.append(part)
                notes.append(
                    f"Auto-split table '{title}' into {total} slides "
                    f"({len(rows)} rows → chunks of ≤{_TABLE_ROWS_SOFT})."
                )
                continue

        # --- Body split (string only) ---
        if isinstance(body, str):
            soft = _BODY_SOFT.get(layout, _BODY_SOFT_DEFAULT)
            # Also match normalized / predefined names.
            soft = _BODY_SOFT.get(_normalize_layout(layout).title().replace(" And ", " + "), soft)
            for key, val in _BODY_SOFT.items():
                if _normalize_layout(key) == _normalize_layout(layout):
                    soft = val
                    break
            parts = _split_body_text(body, soft)
            if len(parts) > 1:
                title = fields.get("title") or "Suite"
                total = len(parts)
                for part_i, chunk in enumerate(parts, start=1):
                    part = copy.deepcopy(slide)
                    # Charts/tables on a body-overflow slide stay on part 1 only.
                    if part_i > 1:
                        part.pop("chart", None)
                        part.pop("table", None)
                        part.pop("image", None)
                        part.pop("images", None)
                        part.pop("image_placeholders", None)
                        # Prefer a plain body layout continuation when possible.
                        if "chart" in _normalize_layout(layout):
                            part["layout"] = "Title + Body"
                    part_fields = dict(part.get("fields") or {})
                    part_fields["title"] = _suffix_title(title, part_i, total)
                    part_fields["body"] = chunk
                    part["fields"] = part_fields
                    new_slides.append(part)
                notes.append(
                    f"Auto-split body '{title}' into {total} slides "
                    f"({len(body)} chars, soft={soft})."
                )
                continue

        new_slides.append(slide)

    out["slides"] = new_slides
    return out, notes


def collect_capacity_findings(
    slides: List[Dict[str, Any]],
    audit_fn,
) -> List[Dict[str, Any]]:
    """Run the existing capacity auditor and return structured findings."""
    findings: List[Dict[str, Any]] = []
    for i, slide_spec in enumerate(slides):
        try:
            raw = audit_fn(i, slide_spec)
        except Exception:
            continue
        layout = slide_spec.get("layout") or "BLANK"
        for severity, message in raw:
            findings.append(
                {
                    "slide": i + 1,
                    "layout": layout,
                    "severity": severity,
                    "message": message,
                }
            )
    return findings


def _looks_risky_image_url(url: str) -> Optional[str]:
    """Return a skip reason for URLs Google Slides cannot fetch, else None.

    Note: public app file routes like `https://wegen…/api/file/<id>` are NOT
    treated as risky — they are anonymously fetchable (no cookies). We still
    HEAD/GET-check them in preflight.
    """
    if not url:
        return "empty url"
    u = url.lower().strip()
    if u.startswith("data:"):
        return "data: URLs are not supported by Slides replaceImage"
    if "localhost" in u or "127.0.0.1" in u:
        return "localhost is unreachable from Google's image fetcher"
    if re.match(r"https?://(10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)", u):
        return "private/LAN IP is unreachable from Google's image fetcher"
    if u.endswith(".svg") or ".svg?" in u:
        return "external SVG is often rejected by Slides"
    return None


async def preflight_image_urls(deck: Dict[str, Any]) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """Drop image URLs that are clearly unfetchable; HEAD/GET-check the rest.

    Mutates a deep copy: bad `image_placeholders` entries and free `image`
    blocks are removed (or the whole field dropped). Reports every decision.
    Public `/api/file/` routes are allowed through when HEAD/GET succeeds.
    """
    out = copy.deepcopy(deck)
    report: List[Dict[str, Any]] = []

    async def _probe_ok(url: str) -> Tuple[bool, str]:
        risky = _looks_risky_image_url(url)
        if risky:
            return False, risky

        def _do_probe() -> Tuple[bool, str]:
            headers = {"User-Agent": "google-workspace-mcp-preflight/1.0"}
            # Prefer HEAD; fall back to GET (some app file routes reject HEAD).
            for method in ("HEAD", "GET"):
                req = urllib.request.Request(url, method=method, headers=headers)
                try:
                    with urllib.request.urlopen(req, timeout=6) as resp:
                        code = getattr(resp, "status", None) or resp.getcode()
                        if code and 200 <= int(code) < 400:
                            return True, f"{method} {code}"
                        if method == "HEAD":
                            continue
                        return False, f"{method} {code}"
                except urllib.error.HTTPError as e:
                    # 401/403/405 on HEAD → try GET; on GET keep URL only if
                    # inconclusive (some CDNs block our UA but allow Google).
                    if method == "HEAD" and e.code in (401, 403, 405, 501):
                        continue
                    if e.code in (401, 403, 405):
                        return True, f"{method} {e.code} (inconclusive — keeping URL)"
                    return False, f"{method} HTTP {e.code}"
                except Exception as e:
                    if method == "HEAD":
                        continue
                    return True, f"probe failed ({type(e).__name__}) — keeping URL"
            return True, "probe inconclusive — keeping URL"

        return await asyncio.to_thread(_do_probe)

    for slide_idx, slide in enumerate(out.get("slides") or []):
        # image_placeholders
        placeholders = slide.get("image_placeholders")
        if isinstance(placeholders, list) and placeholders:
            kept: List[Any] = []
            for item in placeholders:
                url = item if isinstance(item, str) else (item or {}).get("url")
                if not url:
                    continue
                ok, reason = await _probe_ok(str(url))
                entry = {
                    "slide": slide_idx + 1,
                    "kind": "image_placeholder",
                    "url": str(url)[:200],
                    "kept": ok,
                    "reason": reason,
                }
                report.append(entry)
                if ok:
                    kept.append(item)
                else:
                    logger.warning(
                        f"[create_audit_presentation] Image preflight dropped "
                        f"slide #{slide_idx + 1} placeholder URL: {reason} — {url[:120]}"
                    )
            if kept:
                slide["image_placeholders"] = kept
            else:
                slide.pop("image_placeholders", None)

        # free-floating image (singular)
        image = slide.get("image")
        if isinstance(image, dict) and image.get("url"):
            url = str(image["url"])
            ok, reason = await _probe_ok(url)
            report.append(
                {
                    "slide": slide_idx + 1,
                    "kind": "image",
                    "url": url[:200],
                    "kept": ok,
                    "reason": reason,
                }
            )
            if not ok:
                logger.warning(
                    f"[create_audit_presentation] Image preflight dropped "
                    f"slide #{slide_idx + 1} free image: {reason} — {url[:120]}"
                )
                slide.pop("image", None)

        # free-floating images[] (multi)
        images = slide.get("images")
        if isinstance(images, list) and images:
            kept_imgs: List[Any] = []
            for item in images:
                url = item if isinstance(item, str) else (item or {}).get("url")
                if not url:
                    continue
                ok, reason = await _probe_ok(str(url))
                report.append(
                    {
                        "slide": slide_idx + 1,
                        "kind": "images",
                        "url": str(url)[:200],
                        "kept": ok,
                        "reason": reason,
                    }
                )
                if ok:
                    kept_imgs.append(item)
                else:
                    logger.warning(
                        f"[create_audit_presentation] Image preflight dropped "
                        f"slide #{slide_idx + 1} images[] URL: {reason} — {url[:120]}"
                    )
            if kept_imgs:
                slide["images"] = kept_imgs
            else:
                slide.pop("images", None)

    return out, report


async def preprocess_deck(
    deck: Dict[str, Any],
    *,
    presentation_meta: Optional[Dict[str, Any]] = None,
    audit_fn=None,
    run_image_preflight: bool = True,
) -> PreprocessResult:
    """Full preprocess pipeline used by create_audit_presentation."""
    theme = extract_theme(presentation_meta) if presentation_meta else ThemeInfo()
    working = apply_text_defaults(deck, theme=theme)
    working, inferred = infer_chart_positions(working)
    working, split_notes = auto_split_slides(working)

    image_report: List[Dict[str, Any]] = []
    if run_image_preflight:
        working, image_report = await preflight_image_urls(working)

    findings: List[Dict[str, Any]] = []
    if audit_fn is not None:
        findings = collect_capacity_findings(working.get("slides") or [], audit_fn)

    return PreprocessResult(
        deck=working,
        capacity_findings=findings,
        split_notes=split_notes,
        image_preflight=image_report,
        theme_font=theme.font_family,
        inferred_chart_positions=inferred,
    )
