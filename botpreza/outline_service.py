"""Outline/mindmap preview built from slide structure.

Emits a compact Markdown tree of the deck so the user can glance at the
narrative before the PPTX finishes rendering. Pure local logic — no LLM calls,
no network, works from the already-generated slide JSON.
"""

import json
import re


MAX_TITLE = 90
MAX_BULLET = 140
MAX_BULLETS_PER_SLIDE = 4
MAX_OUTLINE_CHARS = 3500  # Telegram message safety


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _get_content(slide: dict) -> dict:
    content = slide.get("content")
    return content if isinstance(content, dict) else {}


def _slide_title(slide: dict, idx: int) -> str:
    content = _get_content(slide)
    for src in (slide, content):
        for key in ("title", "headline", "heading", "slide_title", "name", "topic"):
            value = _clean(src.get(key) or "")
            if value:
                return value[:MAX_TITLE]
    return f"Слайд {idx}"


def _slide_bullets(slide: dict) -> list[str]:
    content = _get_content(slide)
    raw = slide.get("bullets") or content.get("bullets") or []
    if not isinstance(raw, list):
        return []
    out: list[str] = []
    for item in raw:
        if isinstance(item, dict):
            value = _clean(item.get("text") or item.get("label") or item.get("title") or "")
        else:
            value = _clean(item)
        if value:
            out.append(value[:MAX_BULLET])
        if len(out) >= MAX_BULLETS_PER_SLIDE:
            break
    return out


def _slide_layout(slide: dict) -> str:
    layout = _clean(slide.get("layout_type") or _get_content(slide).get("layout_type") or "")
    return layout


def _layout_emoji(layout: str) -> str:
    mapping = {
        "cover": "🧭",
        "agenda": "🗂",
        "context": "🌍",
        "key_findings": "🔑",
        "stat_highlight": "📊",
        "quote_highlight": "💬",
        "comparison": "⚖️",
        "takeaways": "✅",
        "data_matrix": "📈",
        "concept_map": "🧠",
        "process_flow": "🔀",
        "timeline": "🕰",
    }
    return mapping.get(layout, "•")


def build_outline_markdown(formatted_result: str | list | dict, deck_title: str | None = None) -> str:
    if isinstance(formatted_result, str):
        try:
            payload = json.loads(formatted_result)
        except Exception:
            return ""
    else:
        payload = formatted_result

    if isinstance(payload, dict) and isinstance(payload.get("slides"), list):
        slides = payload["slides"]
    elif isinstance(payload, list):
        slides = payload
    else:
        return ""

    if not slides:
        return ""

    lines: list[str] = []
    if deck_title:
        lines.append(f"🧠 *{_clean(deck_title)[:MAX_TITLE]}*")
    lines.append("🗺 Маппинг повествования:")

    for idx, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue
        title = _slide_title(slide, idx)
        layout = _slide_layout(slide)
        emoji = _layout_emoji(layout)
        lines.append(f"{emoji} *{idx}. {title}*")
        for bullet in _slide_bullets(slide):
            lines.append(f"   — {bullet}")

    outline = "\n".join(lines)
    if len(outline) > MAX_OUTLINE_CHARS:
        outline = outline[: MAX_OUTLINE_CHARS - 3].rstrip() + "..."
    return outline
