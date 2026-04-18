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

MAX_INPUT_CHARS = 180000
RETRY_INPUT_CHARS = 45000
MODEL_TIMEOUT_SEC = 90

def _env_int(name: str, default: int) -> int:
    raw = (os.getenv(name) or "").strip()
    try:
        return int(raw) if raw else default
    except ValueError:
        return default

SLIDE_COUNT_MIN = max(4, _env_int("PRESENTATION_MIN_SLIDES", 6))
SLIDE_COUNT_MAX = max(SLIDE_COUNT_MIN, _env_int("PRESENTATION_MAX_SLIDES", 10))
TARGET_SLIDE_COUNT = max(SLIDE_COUNT_MIN, min(SLIDE_COUNT_MAX, _env_int("PRESENTATION_TARGET_SLIDES", 8)))
ULTRA_CHEAP_MODE = os.getenv("ULTRA_CHEAP_MODE", "0").strip().lower() in {"1", "true", "yes", "on"}
DISABLE_TEXT_RETRY = os.getenv("DISABLE_TEXT_RETRY", "0").strip().lower() in {"1", "true", "yes", "on"}
BRAIN_MODEL = os.getenv("ARTEMOX_BRAIN_MODEL", "").strip() or (ARTEMOX_MODEL or "gemini-2.5-pro")
EXTRACTION_MODEL = os.getenv("ARTEMOX_EXTRACTION_MODEL", "").strip() or BRAIN_MODEL
GROUNDING_MODEL = os.getenv("ARTEMOX_GROUNDING_MODEL", "").strip() or EXTRACTION_MODEL
TWO_STAGE_PIPELINE = os.getenv("TWO_STAGE_PIPELINE", "1").strip().lower() in {"1", "true", "yes", "on"}
ENABLE_GROUNDING_CHECK = os.getenv("ENABLE_GROUNDING_CHECK", "1").strip().lower() in {"1", "true", "yes", "on"}

if ULTRA_CHEAP_MODE:
    MAX_INPUT_CHARS = min(MAX_INPUT_CHARS, 24000)
    RETRY_INPUT_CHARS = min(RETRY_INPUT_CHARS, 12000)
    MODEL_TIMEOUT_SEC = min(MODEL_TIMEOUT_SEC, 45)
    SLIDE_COUNT_MAX = min(SLIDE_COUNT_MAX, 8)
    TARGET_SLIDE_COUNT = min(TARGET_SLIDE_COUNT, SLIDE_COUNT_MAX)
    TWO_STAGE_PIPELINE = False
    ENABLE_GROUNDING_CHECK = False

ALLOWED_LAYOUTS = (
    "cover",
    "agenda",
    "context",
    "key_findings",
    "split_infographic",
    "timeline",
    "data_matrix",
    "quote_highlight",
    "stat_highlight",
    "comparison",
    "takeaways",
    "hero_concept",
)

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
    "Ты — Арт-директор и редактор-аналитик уровня NotebookLM. Твоя задача — превратить исходный материал "
    "в связное повествование из структурированных слайдов формата 16:9 для гибридной презентации.\n\n"
    f"{NOTEBOOKLM_PRODUCT_BRIEF}\n\n"
    "Ты НЕ создаёшь текст для PowerPoint вручную. Ты продумываешь структуру повествования, "
    "подробные промпты для генератора картинок (фоновые визуалы) и выдаёшь структурированный JSON, "
    "который программа сама соберёт в слайд.\n\n"
    f"Собери презентацию из {SLIDE_COUNT_MIN}–{SLIDE_COUNT_MAX} слайдов (ориентир {TARGET_SLIDE_COUNT}). "
    "Количество слайдов подбирай под объём и плотность исходного материала: "
    f"меньше материала — ближе к {SLIDE_COUNT_MIN}, плотный и многокомпонентный — ближе к {SLIDE_COUNT_MAX}.\n\n"
    "Повествование должно идти как в NotebookLM Audio Overview: обложка → повестка → контекст/вводная → "
    "ключевые находки с цифрами и цитатами → сравнения/сценарии/процесс → выводы → takeaways.\n\n"
    "Верни строго валидный JSON-массив без markdown и без пояснений. Каждый элемент массива — один слайд "
    "со следующей схемой:\n"
    '{\n'
    '  "slide_number": 1,\n'
    '  "layout_type": "cover|agenda|context|key_findings|split_infographic|timeline|data_matrix|quote_highlight|stat_highlight|comparison|takeaways|hero_concept",\n'
    '  "visual_meta_prompt": "подробный промпт фонового изображения без текста",\n'
    '  "content": {\n'
    '    "title": "краткий, осмысленный заголовок",\n'
    '    "subtitle": "опциональная лид-строка/дескриптор (до 140 симв.)",\n'
    '    "bullets": ["2-5 содержательных буллетов"],\n'
    '    "stats": [ {"value": "42%", "label": "что это значит"} ],\n'
    '    "quotes": [ {"text": "короткая цитата из источника", "author": "кто сказал или источник"} ],\n'
    '    "source_hint": "на каком фрагменте исходника основан слайд (1 предложение)"\n'
    '  },\n'
    '  "speaker_notes": "3-5 предложений для устного сопровождения слайда на уровне аналитика-докладчика"\n'
    '}\n\n'
    "Правила:\n"
    "1. content.title обязателен и содержателен (не «Слайд 2»). content.bullets — 2-5 шт. для наложения на фон.\n"
    "2. Разрешённые layout_type: cover, agenda, context, key_findings, split_infographic, timeline, data_matrix, "
    "quote_highlight, stat_highlight, comparison, takeaways, hero_concept. Выбор типа подчиняй смыслу слайда.\n"
    "3. Обязательная композиция: первый слайд — cover (обложка с названием и идеей), второй — agenda "
    "(повестка на 3-6 пунктов), предпоследний или последний — takeaways (3-5 тезисов-выводов). "
    f"Набор из {SLIDE_COUNT_MIN}+ слайдов обязан включать как минимум один слайд с цифрами (stat_highlight/data_matrix/key_findings со stats) "
    "и хотя бы один quote_highlight/comparison, если материал это позволяет.\n"
    "4. Поля stats и quotes заполняй только если цифры/цитаты реально есть или логично выводятся из материала. "
    "Нельзя выдумывать несуществующие цифры — если точных данных нет, оставь stats/quotes пустыми массивами.\n"
    "5. source_hint — короткая опора на источник (например: «На основе раздела о ключевых факторах роста»). "
    "speaker_notes — разговорные заметки докладчика, раскрывающие слайд, без повторения буллетов дословно.\n"
    "6. visual_meta_prompt — только визуальный фон: композиция, палитра, свет, материалы, атмосфера, глубина. "
    "Фон должен поддерживать layout_type, оставляя чистые зоны под текст/карточки/схемы. "
    "Запрещены: любой текст, слова, буквы на фоне, стоковые люди в офисе, рукопожатия, совещания.\n"
    "7. Ориентир: аналитические editorial slides уровня strategy report / NotebookLM brief — матрицы, схемы, карточки, "
    "техно-диаграммы, график+вывод, архитектурные композиции.\n"
    "8. Между слайдами должно быть явное повествовательное развитие: контекст → находки → сравнения → выводы. "
    "Не допускай дубликатов и пустых слайдов.\n"
)


EXTRACTION_PROMPT = (
    "Ты — аналитик-редактор уровня NotebookLM. Твоя задача — внимательно прочитать исходный материал "
    "и вернуть структурированный 'evidence brief' для последующей сборки презентации.\n\n"
    "Формат — строго валидный JSON без markdown, без пояснений. Схема:\n"
    "{\n"
    '  "topic": "о чём материал в одной фразе",\n'
    '  "audience_guess": "для кого этот материал (1 фраза)",\n'
    '  "thesis": "главный тезис или вывод материала (1-2 предложения)",\n'
    '  "key_facts": ["6-12 ключевых фактов/тезисов, каждый как законченная мысль"],\n'
    '  "numbers": [ {"value": "42%", "label": "что значит", "source_span": "короткая цитата из исходника (до 180 симв.)"} ],\n'
    '  "quotes": [ {"text": "цитата дословно или близко", "author": "кто/источник", "source_span": "контекст (до 180 симв.)"} ],\n'
    '  "entities": ["ключевые имена/бренды/термины"],\n'
    '  "tensions": ["1-4 противоречия или компромисса, если есть"],\n'
    '  "timeline": [ {"when": "2024", "what": "что произошло"} ],\n'
    '  "comparisons": [ {"a": "сторона A", "b": "сторона B", "takeaway": "что важно"} ],\n'
    '  "takeaways": ["3-6 практических выводов/действий"],\n'
    '  "outline": ["линейный порядок тем для презентации (6-10 пунктов)"]\n'
    "}\n\n"
    "Правила:\n"
    "1. Все числа и цитаты должны быть реально извлечены из исходника. source_span — фрагмент текста оригинала, "
    "откуда ты взял факт (1-2 короткие фразы). Если нет — не выдумывай, оставь пустым массивом.\n"
    "2. key_facts формулируй ёмко (до 160 символов), без воды. Это будут якоря для буллетов.\n"
    "3. Не сокращай язык исходника: пиши на том же языке, что и материал (обычно русский).\n"
    "4. Если материал очень короткий или бессодержательный — всё равно верни JSON с пустыми массивами, "
    "но заполненными topic/thesis своими словами.\n"
)


COMPOSITION_PROMPT = (
    "Ты — Арт-директор и редактор-аналитик уровня NotebookLM. У тебя уже есть evidence brief по материалу "
    "(ключевые факты, цифры, цитаты, таймлайн, выводы). Твоя задача — превратить evidence в связную презентацию.\n\n"
    f"{NOTEBOOKLM_PRODUCT_BRIEF}\n\n"
    "ПРАВИЛО ОСНОВАНИЯ: КАЖДАЯ цифра и цитата в слайдах ДОЛЖНА ссылаться на evidence.numbers[] / evidence.quotes[]. "
    "Ничего нельзя выдумывать. Если в evidence нет числа — на слайде числа тоже нет.\n\n"
    f"Собери презентацию из {SLIDE_COUNT_MIN}–{SLIDE_COUNT_MAX} слайдов (ориентир {TARGET_SLIDE_COUNT}). "
    "Количество слайдов подбирай под объём и плотность evidence.\n\n"
    "Повествование как в NotebookLM Audio Overview: cover → agenda → context → ключевые находки с цифрами/цитатами → "
    "сравнения/таймлайн/сценарии → takeaways.\n\n"
    "Верни строго валидный JSON-массив слайдов без markdown, по схеме:\n"
    '[ {\n'
    '  "slide_number": 1,\n'
    '  "layout_type": "cover|agenda|context|key_findings|split_infographic|timeline|data_matrix|quote_highlight|stat_highlight|comparison|takeaways|hero_concept",\n'
    '  "visual_meta_prompt": "подробный промпт фона без текста в изображении",\n'
    '  "content": {\n'
    '    "title": "краткий заголовок",\n'
    '    "subtitle": "опциональный лид",\n'
    '    "bullets": ["2-5 содержательных буллетов"],\n'
    '    "stats": [ {"value": "42%", "label": "…", "source_span": "…"} ],\n'
    '    "quotes": [ {"text": "…", "author": "…", "source_span": "…"} ],\n'
    '    "source_hint": "на каком фрагменте evidence основан слайд (1 предложение)"\n'
    '  },\n'
    '  "speaker_notes": "3-5 предложений разговорных заметок докладчика"\n'
    '} ]\n\n'
    "Правила:\n"
    "1. Первый слайд — cover, второй — agenda (3-6 пунктов из evidence.outline), последний — takeaways.\n"
    "2. Layout подбирай по смыслу слайда: stat_highlight — один акцент-цифра; quote_highlight — пулл-цитата; "
    "comparison — два сценария; timeline — этапы; data_matrix — план/факт; key_findings — 3-4 находки с цифрами.\n"
    "3. Каждое число/цитата содержит source_span из evidence. Если source_span недоступен — не вставляй число/цитату.\n"
    "4. speaker_notes — как докладчик объяснит слайд вслух, без дублирования буллетов дословно.\n"
    "5. visual_meta_prompt — только визуал фона, без текста в изображении. Фон держит чистые зоны под оверлей.\n"
    "6. Язык — как в evidence (обычно русский).\n"
)


GROUNDING_PROMPT = (
    "Проверь презентацию на основание (grounding). Тебе даны evidence brief и draft слайдов. "
    "Верни чистый JSON-массив слайдов в том же формате, но: \n"
    "- удали stats, для которых нет совпадения по value/label в evidence.numbers;\n"
    "- удали quotes, для которых нет совпадения в evidence.quotes;\n"
    "- не меняй структуру, заголовки, bullets и speaker_notes, если они логически опираются на evidence.key_facts;\n"
    "- если после чистки слайд stat_highlight остался без цифры — поменяй layout на split_infographic и оставь bullets;\n"
    "- если quote_highlight остался без цитаты — поменяй layout на context.\n"
    "Ничего не добавляй. Верни только валидный JSON-массив слайдов.\n"
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


def _fallback_stats(units: list[str]) -> list[dict]:
    stats: list[dict] = []
    pattern = re.compile(r"(\d+[\d\s.,]*\s?(?:%|млрд|млн|трлн|тыс|руб|₽|\$|долл|год|мес|дней|человек|раз)?)", re.IGNORECASE)
    seen_values: set[str] = set()
    for unit in units:
        match = pattern.search(unit)
        if not match:
            continue
        value = re.sub(r"\s+", " ", match.group(1)).strip()
        if not value or value in seen_values:
            continue
        label_raw = (unit[: match.start()] + unit[match.end():]).strip(" -—:;,.")
        label = _smart_trim(label_raw, 88) or "Показатель из исходного материала"
        stats.append({"value": value[:24], "label": label})
        seen_values.add(value)
        if len(stats) >= 3:
            break
    return stats


def _fallback_quotes(units: list[str]) -> list[dict]:
    quotes: list[dict] = []
    for unit in units:
        match = re.search(r"[«\"]([^«»\"]{25,200})[»\"]", unit)
        if match:
            quotes.append({"text": match.group(1).strip(), "author": "Источник"})
        if len(quotes) >= 2:
            break
    return quotes


def _fallback_speaker_notes(layout_type: str, title: str, bullets: list[str]) -> str:
    lead = {
        "cover": "Открываем презентацию: обозначаем тему, контекст и для кого материал.",
        "agenda": "Озвучиваем повестку: по пунктам проговариваем, какие блоки рассмотрим и в каком порядке.",
        "context": "Задаём контекст: что происходит в отрасли/ситуации и почему эта тема важна именно сейчас.",
        "key_findings": "Перечисляем ключевые находки, подчёркиваем самые сильные цифры и то, что за ними стоит.",
        "stat_highlight": "Акцентируем внимание на главной цифре: поясняем, как она получена и что она означает.",
        "quote_highlight": "Оттеняем тезис цитатой: коротко комментируем, почему она важна для темы.",
        "split_infographic": "Разбираем блок карточек: поясняем каждую, связывая с общей идеей.",
        "timeline": "Проходим по таймлайну: показываем логику этапов и какие действия они подразумевают.",
        "data_matrix": "Читаем матрицу: сравниваем план и реальность, подсвечиваем разрывы и выводы.",
        "comparison": "Сравниваем две позиции и подчёркиваем, в чём принципиальная разница и что из этого следует.",
        "takeaways": "Сводим выводы: проговариваем главные тезисы, действия и дальнейшие шаги.",
        "hero_concept": "Формулируем ключевой концепт слайда и зачем он нужен для общей истории.",
    }.get(layout_type, "Комментируем слайд своими словами, опираясь на буллеты и контекст.")
    bullet_hint = "; ".join(item for item in bullets[:2]) if bullets else ""
    tail = f" Ключевые тезисы: {bullet_hint}." if bullet_hint else ""
    return _clean_inline_text(f"{lead}{tail} Заголовок слайда: «{title}».", 1000)


def _fallback_slide_pack(text_content: str, style: str, target_total: int) -> list[dict]:
    units = _clean_units(_extract_candidate_lines(text_content) + _extract_sentences(text_content), max_items=20)
    keys = _keywords(text_content, 4)
    topic = _topic_phrase(keys)
    signals = _signal_bullets(units, limit=6)
    stats_pool = _fallback_stats(units)
    quotes_pool = _fallback_quotes(units)

    target_total = max(SLIDE_COUNT_MIN, min(SLIDE_COUNT_MAX, target_total or TARGET_SLIDE_COUNT))
    layout_sequence = _default_layout_sequence(target_total)

    titles_map = {
        "cover": f"{topic}: аналитический обзор",
        "agenda": "Повестка: что обсудим сегодня",
        "context": "Контекст и рамка анализа",
        "key_findings": "Ключевые находки",
        "stat_highlight": "Цифра, определяющая историю",
        "quote_highlight": "Голос источника",
        "split_infographic": "Факторы и акценты",
        "timeline": "Последовательность шагов",
        "data_matrix": "Приоритеты: План vs Реальность",
        "comparison": "Сравнение сценариев",
        "takeaways": "Выводы и следующий шаг",
        "hero_concept": "Концепт и позиционирование",
    }

    bullets_map = {
        "cover": [
            f"Тема: {topic}.",
            "Формат: краткий аналитический разбор на уровне brief.",
            "Аудитория: принимающие решения и команда эксперта.",
        ],
        "agenda": [
            "Контекст и предпосылки обсуждения.",
            "Ключевые находки и цифры.",
            "Сравнения, сценарии и выводы.",
            "Takeaways и следующие шаги.",
        ],
        "context": _fallback_bullets("intro", keys),
        "key_findings": (signals[:3] + _fallback_bullets("middle", keys))[:4],
        "stat_highlight": (signals[:2] + _fallback_bullets("middle", keys))[:3],
        "quote_highlight": (signals[:1] + _fallback_bullets("middle", keys))[:2],
        "split_infographic": (signals[1:4] + _fallback_bullets("middle", keys))[:4],
        "timeline": (signals[:3] + _fallback_bullets("middle", keys))[:4],
        "data_matrix": _fallback_bullets("conclusion", keys),
        "comparison": (signals[:2] + _fallback_bullets("middle", keys))[:2],
        "takeaways": _fallback_bullets("conclusion", keys),
        "hero_concept": _fallback_bullets("intro", keys),
    }

    slides: list[dict] = []
    actual_total = len(layout_sequence)
    for idx, layout_type in enumerate(layout_sequence, 1):
        title = _smart_trim(titles_map.get(layout_type, f"Слайд {idx}"), 88)
        raw_bullets = bullets_map.get(layout_type, _fallback_bullets("middle", keys))
        clean_bullets = _clean_units(raw_bullets, min_len=14, max_items=4)
        if len(clean_bullets) < 3:
            section = "intro" if idx == 1 else "conclusion" if idx == actual_total else "middle"
            clean_bullets = (clean_bullets + _fallback_bullets(section, keys))[:3]

        slide_stats: list[dict] = []
        slide_quotes: list[dict] = []
        if layout_type in ("key_findings", "stat_highlight", "data_matrix") and stats_pool:
            slide_stats = stats_pool[:3]
        if layout_type == "quote_highlight" and quotes_pool:
            slide_quotes = quotes_pool[:1]

        subtitle = ""
        if layout_type == "cover":
            subtitle = "Структурированный разбор материала: контекст, находки, выводы."
        elif layout_type == "agenda":
            subtitle = "Маршрут презентации и логика перехода от контекста к выводам."

        speaker_notes = _fallback_speaker_notes(layout_type, title, clean_bullets)
        source_hint = "На основе автоматически извлечённых тезисов исходного материала."

        content_payload = {
            "title": title,
            "subtitle": subtitle,
            "bullets": clean_bullets[:4],
            "stats": slide_stats,
            "quotes": slide_quotes,
            "source_hint": source_hint,
        }

        slides.append(
            {
                "slide_number": idx,
                "title": title,
                "subtitle": subtitle,
                "bullets": clean_bullets[:4],
                "stats": slide_stats,
                "quotes": slide_quotes,
                "source_hint": source_hint,
                "speaker_notes": speaker_notes,
                "layout_type": layout_type,
                "visual_meta_prompt": _build_image_prompt(
                    title,
                    style,
                    idx,
                    idx == 1,
                    idx == actual_total,
                    layout_type,
                ),
                "content": content_payload,
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


def _default_layout_sequence(total_slides: int) -> list[str]:
    total = max(4, min(12, total_slides or TARGET_SLIDE_COUNT))
    base = ["cover", "agenda", "context", "key_findings"]
    middle_pool = [
        "stat_highlight",
        "split_infographic",
        "timeline",
        "data_matrix",
        "quote_highlight",
        "comparison",
    ]
    tail = ["takeaways"]
    needed_middle = max(0, total - len(base) - len(tail))
    middle: list[str] = []
    for idx in range(needed_middle):
        middle.append(middle_pool[idx % len(middle_pool)])
    return (base + middle + tail)[:total]


def _pick_layout_type(title: str, bullets: list[str], slide_index: int, total_slides: int) -> str:
    text = " ".join([title, *bullets]).lower()
    sequence = _default_layout_sequence(total_slides)
    if 1 <= slide_index <= len(sequence):
        default_choice = sequence[slide_index - 1]
    else:
        default_choice = "split_infographic"

    if slide_index == 1:
        return "cover"
    if slide_index == total_slides:
        return "takeaways"
    if slide_index == 2 and total_slides >= 4:
        return "agenda"

    if any(token in text for token in ("цитат", "quote", "по словам", "писал:", "отметил")):
        return "quote_highlight"
    if any(token in text for token in ("%", "млрд", "млн", "трлн", "рост", "прирост", "доля", "kpi", "метрик")):
        return "stat_highlight"
    if any(token in text for token in ("план", "факт", "реальность", "сравн", "trade-off", "vs", "квадрант", "матриц")):
        return "data_matrix"
    if any(token in text for token in ("этап", "процесс", "метод", "pipeline", "stack", "слой", "уров", "roadmap", "последовательн")):
        return "timeline"
    if any(token in text for token in ("контекст", "рамка", "парадигма", "концепт", "манифест", "позиционирование")):
        return "context"
    if any(token in text for token in ("вывод", "итог", "takeaway", "следующ")):
        return "takeaways"
    return default_choice


def _layout_hint(layout_type: str) -> str:
    hints = {
        "cover": "monumental editorial title scene, cinematic depth, calm central focal area for a bold headline, premium publication cover feel",
        "agenda": "subtle architectural grid evoking a table of contents, soft verticals, calm breathing space for a clean numbered list overlay",
        "context": "atmospheric establishing shot that sets the thematic stage, wide horizon, soft lighting, space for a contextual narrative block",
        "key_findings": "analytical editorial composition with clean zones for large numbers, side notes and annotation ribbons",
        "stat_highlight": "bold hero composition built around a single commanding numeric focal point, spotlight lighting, minimal supporting texture",
        "quote_highlight": "quiet editorial spread with a soft portrait-like ambiance and a generous quiet zone reserved for a pulled quote",
        "comparison": "symmetric split composition with two balanced halves ready for side-by-side contrast blocks",
        "takeaways": "closing editorial spread with a summary grid feel, warm finishing tones, clean list area on one side",
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


def _normalize_model_bullets(raw_value, *, limit: int = 5) -> list[str]:
    if isinstance(raw_value, list):
        values = raw_value
    else:
        values = re.split(r"[\n\r]+", str(raw_value or ""))
    bullets = []
    for item in values:
        cleaned = _clean_bullet(str(item or ""))
        if cleaned:
            bullets.append(cleaned)
    return bullets[:limit]


def _clean_inline_text(value, limit: int) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip().strip("«»\"' ")
    return text[:limit]


def _normalize_stats(raw_value) -> list[dict]:
    if not isinstance(raw_value, list):
        return []
    stats: list[dict] = []
    for item in raw_value:
        if isinstance(item, dict):
            value = _clean_inline_text(item.get("value") or item.get("number") or item.get("metric"), 24)
            label = _clean_inline_text(item.get("label") or item.get("caption") or item.get("description"), 90)
        else:
            text = _clean_inline_text(item, 110)
            match = re.match(r"(\S+?)\s*[-—:]\s*(.+)", text)
            if match:
                value = _clean_inline_text(match.group(1), 24)
                label = _clean_inline_text(match.group(2), 90)
            else:
                value = ""
                label = text
        if value or label:
            stats.append({"value": value, "label": label})
        if len(stats) >= 4:
            break
    return stats


def _normalize_quotes(raw_value) -> list[dict]:
    if not isinstance(raw_value, list):
        return []
    quotes: list[dict] = []
    for item in raw_value:
        if isinstance(item, dict):
            text = _clean_inline_text(item.get("text") or item.get("quote") or item.get("body"), 220)
            author = _clean_inline_text(item.get("author") or item.get("source") or item.get("by"), 90)
        else:
            text = _clean_inline_text(item, 220)
            author = ""
        if text:
            quotes.append({"text": text, "author": author})
        if len(quotes) >= 3:
            break
    return quotes


def _normalize_speaker_notes(raw_value) -> str:
    if isinstance(raw_value, list):
        joined = " ".join(str(item or "").strip() for item in raw_value if str(item or "").strip())
    else:
        joined = str(raw_value or "").strip()
    joined = re.sub(r"\s+", " ", joined).strip()
    return joined[:1200]


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
    if len(slides) < SLIDE_COUNT_MIN:
        return ""

    actual_total = max(SLIDE_COUNT_MIN, min(SLIDE_COUNT_MAX, len(slides)))
    slides = slides[:actual_total]

    normalized: list[dict] = []
    for idx, slide in enumerate(slides, 1):
        content = slide.get("content") if isinstance(slide.get("content"), dict) else {}
        raw_title = (
            content.get("title")
            or content.get("headline")
            or content.get("heading")
            or slide.get("title")
            or slide.get("headline")
            or slide.get("heading")
            or f"Слайд {idx}"
        )
        title = _clean_bullet(str(raw_title))[:90]
        subtitle = _clean_inline_text(
            content.get("subtitle") or content.get("lead") or slide.get("subtitle") or slide.get("lead"),
            180,
        )
        bullets = _normalize_model_bullets(
            content.get("bullets")
            or content.get("theses")
            or content.get("points")
            or slide.get("bullets")
            or slide.get("theses")
            or slide.get("points")
            or slide.get("content"),
            limit=5,
        )
        stats = _normalize_stats(content.get("stats") or slide.get("stats") or content.get("numbers") or slide.get("numbers"))
        quotes = _normalize_quotes(content.get("quotes") or slide.get("quotes") or content.get("citations") or slide.get("citations"))
        source_hint = _clean_inline_text(
            content.get("source_hint")
            or content.get("source")
            or slide.get("source_hint")
            or slide.get("source"),
            220,
        )
        speaker_notes = _normalize_speaker_notes(
            slide.get("speaker_notes")
            or slide.get("notes")
            or content.get("speaker_notes")
            or content.get("notes")
        )

        layout_raw = str(slide.get("layout_type") or content.get("layout_type") or "").strip().lower()
        layout_type = layout_raw if layout_raw in ALLOWED_LAYOUTS else _pick_layout_type(title, bullets, idx, actual_total)
        if idx == 1 and layout_type not in ("cover", "hero_concept"):
            layout_type = "cover"
        if idx == actual_total and layout_type not in ("takeaways", "data_matrix"):
            layout_type = "takeaways"

        visual_meta_prompt = str(
            slide.get("visual_meta_prompt") or slide.get("image_prompt") or ""
        ).strip() or _build_image_prompt(
            title, style, idx, idx == 1, idx == actual_total, layout_type
        )

        content_payload = {
            "title": title or f"Слайд {idx}",
            "subtitle": subtitle,
            "bullets": bullets[:5],
            "stats": stats,
            "quotes": quotes,
            "source_hint": source_hint,
        }

        normalized.append(
            {
                "slide_number": idx,
                "title": title or f"Слайд {idx}",
                "subtitle": subtitle,
                "bullets": bullets[:5],
                "stats": stats,
                "quotes": quotes,
                "source_hint": source_hint,
                "speaker_notes": speaker_notes,
                "layout_type": layout_type,
                "visual_meta_prompt": visual_meta_prompt,
                "image_prompt": visual_meta_prompt,
                "content": content_payload,
            }
        )

    return json.dumps(normalized, ensure_ascii=False)


def _fallback_structure(text_content: str, style: str = "auto") -> str:
    text_len = len((text_content or "").strip())
    if text_len < 1200:
        target_total = SLIDE_COUNT_MIN
    elif text_len > 8000:
        target_total = SLIDE_COUNT_MAX
    else:
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
