import asyncio
import json
import os
import re
from collections import Counter

from artemox_client import (
    ARTEMOX_MODEL,
    ARTEMOX_QUOTA_EXCEEDED_SENTINEL,
    get_artemox_client,
)
from reference_style_service import get_reference_style_brief

MAX_INPUT_CHARS = 45000
RETRY_INPUT_CHARS = 15000
MODEL_TIMEOUT_SEC = 45
TARGET_SLIDE_COUNT = max(3, min(5, int(os.getenv("PRESENTATION_TARGET_SLIDES", "3") or "3")))
ULTRA_CHEAP_MODE = os.getenv("ULTRA_CHEAP_MODE", "0").strip().lower() in {"1", "true", "yes", "on"}
DISABLE_TEXT_RETRY = os.getenv("DISABLE_TEXT_RETRY", "0").strip().lower() in {"1", "true", "yes", "on"}
BRAIN_MODEL = os.getenv("ARTEMOX_BRAIN_MODEL", "gemini-1.5-pro").strip() or ARTEMOX_MODEL

if ULTRA_CHEAP_MODE:
    MAX_INPUT_CHARS = min(MAX_INPUT_CHARS, 12000)
    RETRY_INPUT_CHARS = min(RETRY_INPUT_CHARS, 8000)
    MODEL_TIMEOUT_SEC = min(MODEL_TIMEOUT_SEC, 30)

STYLE_INSTRUCTIONS = {
    "academic": (
        "Визуальный стиль: строгий академический / корпоративный отчёт. "
        "Доминируют тёмно-синий, белый, серый. Чёткая типографика, инфографика, "
        "диаграммы и цифры в центре внимания. Фоны — абстрактная архитектура, геометрия, свет."
    ),
    "pitch": (
        "Визуальный стиль: питч стартапа. "
        "Яркие контрастные цвета, крупные заголовки, фотореалистичные 3D-объекты, "
        "минимум текста, максимум визуального удара. Каждый слайд = одна мощная мысль."
    ),
    "creative": (
        "Визуальный стиль: креативный / арт-журнал. "
        "Нестандартные ракурсы, коллажная эстетика, сочные цвета, асимметрия. "
        "Смешивай фото, иллюстрации, абстрактные 3D-элементы."
    ),
    "auto": (
        "Визуальный стиль: на усмотрение ИИ. "
        "Сам определи идеальный стиль под тему. Подбери палитру и настроение."
    ),
}

STOPWORDS = {
    "это", "как", "для", "или", "при", "что", "его", "её", "она", "они", "мы", "вы", "над", "под", "из", "по",
    "в", "на", "к", "от", "до", "за", "не", "но", "а", "и", "о", "об", "у", "же", "ли", "бы", "с", "со",
    "the", "and", "for", "with", "from", "that", "this", "into", "about", "over", "under", "are", "was", "were",
    "при", "также", "вместе", "тем", "уровне", "данной", "сферы", "россии", "москвы", "федеральное", "государственное",
    "учреждение", "высшего", "образования", "университет", "правительстве", "российской", "федерации",
    "факультет", "кафедра", "институт", "школа", "управления",
}

NOTEBOOKLM_PRODUCT_BRIEF = (
    "Продуктовая рамка системы:\n"
    "1. Абсолютная всеядность. Система должна уверенно работать с любым сырьём: длинные DOCX и PDF, аудио и голосовые, "
    "фото конспектов и скриншоты графиков, ссылки на статьи и лекции. Из хаотичного материала нужно извлекать факты, "
    "смыслы, структуру, числа, тезисы, аргументы и превращать всё это в ясный storytelling.\n"
    "2. Глубокое понимание контекста. Это не конвертер текста в слайды, а умный редактор. "
    "Нужно учитывать аудиторию, степень формальности, жанр и цель: защита отчёта, доклад, питч, аналитическая записка. "
    "Система сама выбирает, где нужен data-driven подход с матрицами, графиками и сравнениями, "
    "а где уместнее минимализм, крупная метафора или архитектурная композиция.\n"
    "3. Премиальная гибридная верстка. Слайд собирается по принципу 'внутренности и оболочка': "
    "генеративная модель рисует дорогой визуальный фон с чистым негативным пространством, "
    "а программная сборка накладывает строгую типографику, контейнеры, карточки, схемы и акценты. "
    "Итог должен выглядеть как работа команды дизайнеров и редакторов, а не как обычный PowerPoint-шаблон."
)

SYSTEM_PROMPT = (
    "Ты — Арт-директор мирового уровня. Твоя задача — превратить текстовый материал "
    "в серию сильных аналитических слайдов формата 16:9 для гибридной презентации.\n\n"
    f"{NOTEBOOKLM_PRODUCT_BRIEF}\n\n"
    "Ты НЕ создаёшь текст для PowerPoint. Ты придумываешь подробные промпты для "
    "генератора картинок, который рисует только фоновые визуалы, а текст будет наложен отдельно в PowerPoint.\n\n"
    f"Создай ровно {TARGET_SLIDE_COUNT} слайда. "
    "Каждая ключевая мысль, факт, цифра или аргумент заслуживают отдельный слайд.\n\n"
    "Верни строго валидный JSON-массив без markdown и без пояснений. Формат:\n"
    '[{"slide_number": 1, "layout_type": "...", "visual_meta_prompt": "...", "content": {"title": "...", "bullets": ["...", "..."]}}, {"slide_number": 2, "layout_type": "...", "visual_meta_prompt": "...", "content": {"title": "...", "bullets": ["...", "..."]}}]\n\n'
    "Правила:\n"
    "1. Для каждого слайда верни content.title и 2-4 content.bullets для наложения поверх фона.\n"
    "2. layout_type обязателен. Разрешённые типы: split_infographic, data_matrix, hero_concept, timeline.\n"
    "3. visual_meta_prompt описывает только фон слайда, без текста и без слов в изображении.\n"
    "4. Фон должен поддерживать конкретный layout_type: оставляй чистые зоны там, где будет текст, карточки, матрица или схема.\n"
    "5. Укажи композицию, палитру, свет, материалы, атмосферу и глубину.\n"
    "6. Запрещены: любой текст, слова, буквы на фоне, стоковые люди в офисе, рукопожатия, совещания.\n"
    "7. Ориентируйся не на банальный pitch deck, а на аналитические editorial slides уровня strategy report: матрицы, схемы, карточки, техно-диаграммы, график+вывод, архитектурные композиции.\n"
    f"8. На наборе из {TARGET_SLIDE_COUNT} слайдов обязательно должны присутствовать: hero_concept, минимум один timeline/split_infographic, минимум один data_matrix.\n"
    "9. Первый слайд — hero_concept. Последний — data_matrix или split_infographic с итоговым выводом.\n"
    "10. При выборе layout_type учитывай не только текст, но и формат материала: если есть факторы и сравнение план/факт — data_matrix; "
    "если есть последовательность этапов — timeline; "
    "если нужен сильный концептуальный старт — hero_concept; "
    "в остальных случаях — split_infographic.\n"
)


def _prepare_text_for_model(text_content: str, limit: int) -> str:
    normalized = re.sub(r"\s+", " ", (text_content or "")).strip()
    if len(normalized) <= limit:
        return normalized

    head_size = int(limit * 0.7)
    tail_size = limit - head_size
    return (
        normalized[:head_size]
        + "\n\n[... document truncated for analysis ...]\n\n"
        + normalized[-tail_size:]
    )


def _clean_model_response(text: str) -> str:
    cleaned = (text or "").strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.removeprefix("```json").removeprefix("```").strip()
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3].strip()
    return cleaned


def _extract_response_text(response) -> str:
    text = _clean_model_response(getattr(response, "text", ""))
    if text:
        return text

    print(f"Artemox returned empty text. Response object: {response}")
    return ""


def _run_model(prepared_text: str, style: str = "auto") -> str:
    client = get_artemox_client()
    if client is None:
        return ""

    style_block = STYLE_INSTRUCTIONS.get(style, STYLE_INSTRUCTIONS["auto"])
    reference_brief = get_reference_style_brief()

    prompt = (
        f"{SYSTEM_PROMPT}\n\n"
        f"{style_block}\n\n"
        f"Визуальный ориентир по референсным презентациям:\n{reference_brief}\n\n"
        f"Исходный материал для анализа:\n{prepared_text}"
    )
    response = client.models.generate_content(
        model=BRAIN_MODEL,
        contents=prompt,
    )
    return _extract_response_text(response)


def _clean_bullet(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", (text or "").strip())
    cleaned = cleaned.lstrip("-• ").strip(" .")
    return cleaned[:140]


def _normalize_compare_key(text: str) -> str:
    base = re.sub(r"[^a-zа-яё0-9 ]+", " ", (text or "").lower(), flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", base).strip()


def _smart_trim(text: str, limit: int = 120) -> str:
    normalized = re.sub(r"\s+", " ", (text or "").strip())
    normalized = re.sub(r"\(([^)]{1,90})\)\s*\(\1\)", r"(\1)", normalized)
    normalized = re.sub(r"(.{18,80})\s+\1", r"\1", normalized, flags=re.IGNORECASE)
    normalized = normalized.strip(" .;:-")
    if len(normalized) <= limit:
        return normalized
    cut = normalized[:limit]
    for sep in (". ", "; ", ": ", ", ", " "):
        idx = cut.rfind(sep)
        if idx > int(limit * 0.6):
            cut = cut[:idx]
            break
    return cut.strip(" .;:-")


def _clean_units(units: list[str], *, min_len: int = 18, max_items: int = 16) -> list[str]:
    prepared: list[str] = []
    seen: list[str] = []
    banned_fragments = (
        "федеральное государственное образовательное",
        "учреждение высшего образования",
        "при правительстве российской федерации",
        "финансовый университет",
        "факультет",
        "кафедра",
        "институт",
        "высшая школа управления",
    )
    for raw in units:
        text = _smart_trim(_clean_bullet(raw), 130)
        if len(text) < min_len:
            continue
        if text.startswith("(") and text.endswith(")"):
            continue
        lowered = text.lower()
        if any(fragment in lowered for fragment in banned_fragments):
            continue
        key = _normalize_compare_key(text)
        if not key or key in seen:
            continue
        if any((key in prev or prev in key) and min(len(key), len(prev)) > 24 for prev in seen):
            continue
        prepared.append(text)
        seen.append(key)
        if len(prepared) >= max_items:
            break
    return prepared


def _keywords(text: str, limit: int = 4) -> list[str]:
    words = re.findall(r"[A-Za-zА-Яа-яЁё]{4,}", text or "")
    normalized = [w.lower() for w in words if w.lower() not in STOPWORDS]
    if not normalized:
        return []
    freq = Counter(normalized)
    return [word for word, _ in freq.most_common(limit)]


def _title_from_units(units: list[str], text_content: str) -> str:
    if units:
        candidate = _smart_trim(units[0], 62)
        if len(candidate) >= 16:
            return candidate
    keys = _keywords(text_content, 2)
    if len(keys) == 2:
        return f"Аналитика: {keys[0].title()} и {keys[1].title()}"
    if len(keys) == 1:
        return f"Аналитика: {keys[0].title()}"
    return "Ключевые выводы анализа"


def _fallback_bullets(section: str, keys: list[str]) -> list[str]:
    topic = ", ".join(keys[:3]) if keys else "исходные данные и ограничения"
    if section == "intro":
        return [
            "Контекст и задача: структурировать исходный материал в понятную управленческую логику.",
            "Фокус: выделить ключевые факторы и приоритеты принятия решений.",
            "Подход: структурировать материал в формат, удобный для защиты и обсуждения.",
        ]
    if section == "middle":
        return [
            "Сформирована рабочая структура факторов и взаимосвязей между блоками.",
            "Определены узкие места и точки роста по операционной и стратегической осям.",
            "Собраны тезисы для визуального сравнения сценариев и альтернатив.",
        ]
    return [
        "Вывод: приоритеты требуют синхронизации стратегии, процессов и метрик.",
        "Риск: без единой рамки принятия решений возрастает операционная сложность.",
        "Следующий шаг: зафиксировать дорожную карту с контрольными индикаторами.",
    ]


def _topic_phrase(keys: list[str]) -> str:
    if len(keys) >= 2:
        return f"{keys[0].title()} и {keys[1].title()}"
    if len(keys) == 1:
        return keys[0].title()
    return "Системный анализ"


def _signal_bullets(units: list[str], *, limit: int = 3) -> list[str]:
    prepared: list[str] = []
    banned_fragments = (
        "федеральное государственное образовательное",
        "учреждение высшего образования",
        "при правительстве российской федерации",
    )
    for line in units:
        item = _smart_trim(line, 92)
        if len(item) < 18:
            continue
        lowered = item.lower()
        if any(fragment in lowered for fragment in banned_fragments):
            continue
        prepared.append(item)
        if len(prepared) >= limit:
            break
    return prepared


def _fallback_slide_pack(text_content: str, style: str, target_total: int) -> list[dict]:
    units = _clean_units(_extract_candidate_lines(text_content) + _extract_sentences(text_content), max_items=12)
    keys = _keywords(text_content, 4)
    topic = _topic_phrase(keys)
    signals = _signal_bullets(units, limit=3)

    specs = [
        ("hero_concept", "Контекст и рамка анализа", _fallback_bullets("intro", keys)),
        ("timeline", "Ключевые факторы и логика реализации", (signals[:2] + _fallback_bullets("middle", keys))[:3]),
        ("data_matrix", "Приоритеты: План vs Реальность", _fallback_bullets("conclusion", keys)),
        ("split_infographic", "Динамика и контроль исполнения", (signals[1:3] + _fallback_bullets("middle", keys))[:3]),
        ("data_matrix", "Сценарии и итоговый выбор", _fallback_bullets("conclusion", keys)),
    ]
    specs = specs[:max(1, target_total)]

    slides: list[dict] = []
    actual_total = len(specs)
    for idx, (layout_type, title, bullets) in enumerate(specs, 1):
        clean_bullets = _clean_units(bullets, min_len=14, max_items=3)
        if len(clean_bullets) < 3:
            section = "intro" if idx == 1 else "conclusion" if idx == actual_total else "middle"
            clean_bullets = (clean_bullets + _fallback_bullets(section, keys))[:3]
        slides.append(
            {
                "slide_number": idx,
                "title": _smart_trim(title, 74),
                "bullets": clean_bullets[:3],
                "layout_type": layout_type,
                "visual_meta_prompt": _build_image_prompt(
                    title,
                    style,
                    idx,
                    idx == 1,
                    idx == actual_total,
                    layout_type,
                ),
                "content": {
                    "title": _smart_trim(title, 74),
                    "bullets": clean_bullets[:3],
                },
            }
        )
    return slides


def _extract_candidate_lines(text_content: str) -> list[str]:
    lines = []
    for raw in re.split(r"[\n\r]+", text_content):
        cleaned = _clean_bullet(raw)
        if cleaned and len(cleaned) > 8:
            lines.append(cleaned)
    return lines


def _extract_sentences(text_content: str) -> list[str]:
    compact = re.sub(r"\s+", " ", text_content or "").strip()
    parts = re.split(r"(?<=[.!?])\s+", compact)
    result = []
    for part in parts:
        cleaned = _clean_bullet(part)
        if cleaned and len(cleaned) > 20:
            result.append(cleaned)
    return result


def _style_hint_for_prompt(style: str) -> str:
    if style == "academic":
        return "deep navy, graphite, silver palette, architectural abstractions, subtle data-grid motifs"
    if style == "pitch":
        return "high-contrast corporate palette, luminous accents, dynamic 3D forms, bold cinematic depth"
    if style == "creative":
        return "editorial collage energy, expressive gradients, asymmetry, tactile abstract forms"
    return "premium corporate editorial palette, clean geometry, cinematic lighting"


def _pick_layout_type(title: str, bullets: list[str], slide_index: int, total_slides: int) -> str:
    text = " ".join([title, *bullets]).lower()
    sequence_map = {
        3: ["hero_concept", "timeline", "data_matrix"],
        4: ["hero_concept", "split_infographic", "timeline", "data_matrix"],
        5: ["hero_concept", "split_infographic", "timeline", "data_matrix", "split_infographic"],
    }
    if total_slides in sequence_map and 1 <= slide_index <= len(sequence_map[total_slides]):
        return sequence_map[total_slides][slide_index - 1]

    if slide_index == 1:
        return "hero_concept"
    if slide_index == total_slides:
        return "data_matrix"
    if any(token in text for token in ("план", "факт", "реальность", "сравн", "trade-off", "vs", "квадрант", "матриц")):
        return "data_matrix"
    if any(token in text for token in ("этап", "процесс", "метод", "pipeline", "stack", "слой", "уров")):
        return "timeline"
    if any(token in text for token in ("контекст", "рамка", "парадигма", "концепт", "манифест", "позиционирование")):
        return "hero_concept"
    return "split_infographic"


def _layout_hint(layout_type: str) -> str:
    hints = {
        "hero_concept": "strong editorial hero composition with abstract conceptual centerpiece on the right and calm left area",
        "split_infographic": "clean right-side infographic zones for cards and analytical callouts, balanced negative space",
        "timeline": "directional timeline rhythm with sequential anchors on right side, structured flow lines",
        "data_matrix": "clear matrix/grid logic with comparative zones prepared for Plan vs Reality reading",
    }
    return hints.get(layout_type, "clean editorial analytical slide layout")


def _build_image_prompt(topic: str, style: str, slide_index: int, is_title: bool, is_final: bool, layout_type: str) -> str:
    role = "title slide background" if is_title else "closing slide background" if is_final else "section slide background"
    topic_hint = _clean_bullet(topic) or f"slide {slide_index}"
    style_hint = _style_hint_for_prompt(style)
    layout_hint = _layout_hint(layout_type)
    return (
        f"{role}, premium analytical background about {topic_hint}, "
        f"{style_hint}, {layout_hint}, "
        "abstract architectural minimalism, soft urban gradients, premium financial report aesthetic, blurred textures, "
        "space on the left is reserved for text overlay, clean negative space, corporate colors, "
        "soft lighting, 16:9 ratio, no text, no letters, no words"
    )


def _normalize_model_bullets(raw_value) -> list[str]:
    if isinstance(raw_value, list):
        values = raw_value
    else:
        values = re.split(r"[\n\r]+", str(raw_value or ""))
    bullets = []
    for item in values:
        cleaned = _clean_bullet(str(item or ""))
        if cleaned:
            bullets.append(cleaned)
    return bullets[:4]


def _normalize_model_structure(result_text: str, style: str) -> str:
    try:
        payload = json.loads(_clean_model_response(result_text))
    except Exception:
        return ""

    if isinstance(payload, dict):
        slides = payload.get("slides") if isinstance(payload.get("slides"), list) else [payload]
    elif isinstance(payload, list):
        slides = payload
    else:
        return ""

    slides = [slide for slide in slides if isinstance(slide, dict)]
    if len(slides) < TARGET_SLIDE_COUNT:
        return ""

    normalized: list[dict] = []
    actual_total = min(TARGET_SLIDE_COUNT, len(slides))
    for idx, slide in enumerate(slides[:actual_total], 1):
        content = slide.get("content") if isinstance(slide.get("content"), dict) else {}
        title = _clean_bullet(
            str(
                content.get("title")
                or content.get("headline")
                or content.get("heading")
                or slide.get("title")
                or slide.get("headline")
                or slide.get("heading")
                or f"Слайд {idx}"
            )
        )[:80]
        bullets = _normalize_model_bullets(
            content.get("bullets")
            or content.get("theses")
            or content.get("points")
            or slide.get("bullets")
            or slide.get("theses")
            or slide.get("points")
            or slide.get("content")
        )
        layout_type = _pick_layout_type(title, bullets, idx, actual_total)
        visual_meta_prompt = str(slide.get("visual_meta_prompt") or slide.get("image_prompt") or "").strip() or _build_image_prompt(
            title, style, idx, idx == 1, idx == actual_total, layout_type
        )
        normalized.append(
            {
                "slide_number": idx,
                "title": title or f"Слайд {idx}",
                "bullets": bullets[:4],
                "layout_type": layout_type,
                "visual_meta_prompt": visual_meta_prompt,
                "image_prompt": visual_meta_prompt,
                "content": {
                    "title": title or f"Слайд {idx}",
                    "bullets": bullets[:4],
                },
            }
        )

    return json.dumps(normalized, ensure_ascii=False)


def _fallback_structure(text_content: str, style: str = "auto") -> str:
    target_total = TARGET_SLIDE_COUNT
    slides = _fallback_slide_pack(text_content, style, target_total)
    return json.dumps(slides, ensure_ascii=False)


def _looks_like_gateway_block(exc: Exception) -> bool:
    message = str(exc).lower()
    markers = (
        "403",
        "forbidden",
        "cloudflare",
        "challenge",
        "peer closed connection",
        "incomplete chunked read",
        "timed out",
        "timeout",
    )
    return any(marker in message for marker in markers)


def _looks_like_auth_block(exc: Exception) -> bool:
    message = str(exc).lower()
    markers = (
        "401",
        "auth_error",
        "authentication error",
        "key is blocked",
        "blocked. update via `/key/unblock`",
        "unauthorized",
    )
    return any(marker in message for marker in markers)


async def _run_model_with_timeout(prepared_text: str, style: str) -> str:
    return await asyncio.wait_for(
        asyncio.to_thread(_run_model, prepared_text, style),
        timeout=MODEL_TIMEOUT_SEC,
    )


async def analyze_and_create_structure(text_content: str, style: str = "auto") -> str:
    if get_artemox_client() is None or not (text_content or "").strip():
        return _fallback_structure(text_content, style)

    prepared_text = _prepare_text_for_model(text_content, MAX_INPUT_CHARS)
    print(f"Artemox input length: original={len(text_content)}, prepared={len(prepared_text)}")

    try:
        result = await _run_model_with_timeout(prepared_text, style)
        if result:
            normalized = _normalize_model_structure(result, style)
            if normalized:
                return normalized
    except Exception as exc:
        print(f"Artemox primary request failed: {exc}")
        if "429" in str(exc) or "quota" in str(exc).lower() or "limit" in str(exc).lower():
            return ARTEMOX_QUOTA_EXCEEDED_SENTINEL
        if _looks_like_gateway_block(exc) or _looks_like_auth_block(exc):
            return _fallback_structure(text_content, style)

    if DISABLE_TEXT_RETRY or ULTRA_CHEAP_MODE:
        return _fallback_structure(text_content, style)

    retry_text = _prepare_text_for_model(text_content, RETRY_INPUT_CHARS)
    print(f"Artemox retry with reduced input length={len(retry_text)}")
    try:
        result = await _run_model_with_timeout(retry_text, style)
        if result:
            normalized = _normalize_model_structure(result, style)
            if normalized:
                return normalized
    except Exception as exc:
        print(f"Artemox retry request failed: {exc}")
        if "429" in str(exc) or "quota" in str(exc).lower() or "limit" in str(exc).lower():
            return ARTEMOX_QUOTA_EXCEEDED_SENTINEL

    return _fallback_structure(text_content, style)
