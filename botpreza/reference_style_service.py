import json
import os
import shutil
import tempfile
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from PIL import Image

from artemox_client import ARTEMOX_MEDIA_MODEL, get_artemox_client


DEFAULT_REFERENCE_PRESENTATIONS = [
    "/opt/prezbot/references/презентация .pptx",
    "/opt/prezbot/references/Chechnya_Strategic_Audit.pptx",
    "/opt/prezbot/references/Dagestan_Strategy_Versus_Reality_(4).pptx",
    "/opt/prezbot/references/Infrastructure_for_Life_2030.pptx",
    "/opt/prezbot/references/Russian_Pharma_Market_Transformation.pptx",
    "/opt/prezbot/references/Russian_Pharmaceutical_Blueprint_(3).pptx",
    "/opt/prezbot/references/The_Q-cumber_Revolution.pptx",
    "/opt/prezbot/references/The_Fifth_Paradigm.pptx",
    "/opt/prezbot/references/The_Silicon_Evolution.pptx",
    "/opt/prezbot/references/Tula_Innovation_Blueprint.pptx",
]

CACHE_PATH = Path("temp/reference_style_brief.json")
MAX_REFERENCE_DECKS = 6
ULTRA_CHEAP_MODE = os.getenv("ULTRA_CHEAP_MODE", "0").strip().lower() in {"1", "true", "yes", "on"}
DISABLE_REFERENCE_MODEL_ANALYSIS = os.getenv("DISABLE_REFERENCE_MODEL_ANALYSIS", "0").strip().lower() in {"1", "true", "yes", "on"}

FALLBACK_STYLE_BRIEF = (
    "Ориентир по уровню референсов: премиальный editorial/corporate журнал. "
    "Слайды должны выглядеть дорого и собранно: сильная иерархия, много воздуха, "
    "минимум перегруза, 1 доминирующая идея на слайд, аккуратные корпоративные палитры, "
    "контрастные фоновые сцены, абстрактные или метафорические визуалы вместо клише. "
    "Тезисы короткие, заголовки ёмкие, композиция чистая. "
    "Повторяющиеся архетипы из референсов: hero + техническая иллюстрация, 2x2 карточки, график + аналитический вывод, "
    "матрица/квадрант, радиальная анатомия, стек или процесс, карта/кластеры. "
    "Игнорируй любые watermark/лейблы вроде NotebookLM."
)


def _mean_luminance(image: Image.Image) -> float:
    gray = image.convert("L").resize((64, 64))
    values = list(gray.getdata())
    return sum(values) / max(len(values), 1)


def _mean_saturation(image: Image.Image) -> float:
    hsv = image.convert("HSV").resize((64, 64))
    s_values = [pixel[1] for pixel in hsv.getdata()]
    return sum(s_values) / max(len(s_values), 1)


def _local_reference_brief(paths: list[str]) -> str:
    decks = 0
    slide_counts: list[int] = []
    media_counts: list[int] = []
    luminance_values: list[float] = []
    saturation_values: list[float] = []

    for raw_path in paths[:MAX_REFERENCE_DECKS]:
        path = Path(raw_path)
        if not path.exists():
            continue
        try:
            with ZipFile(path) as archive:
                slide_names = sorted(
                    name for name in archive.namelist()
                    if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                )
                media_names = sorted(
                    name for name in archive.namelist()
                    if name.startswith("ppt/media/") and name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))
                )
                if not slide_names:
                    continue
                decks += 1
                slide_counts.append(len(slide_names))
                media_counts.append(len(media_names))
                for media_name in media_names[:2]:
                    try:
                        data = archive.read(media_name)
                        with Image.open(BytesIO(data)) as image:
                            luminance_values.append(_mean_luminance(image))
                            saturation_values.append(_mean_saturation(image))
                    except Exception:
                        continue
        except Exception:
            continue

    if decks == 0:
        return FALLBACK_STYLE_BRIEF

    avg_slides = sum(slide_counts) / max(len(slide_counts), 1)
    avg_media = sum(media_counts) / max(len(media_counts), 1)
    media_ratio = avg_media / max(avg_slides, 1)
    avg_luma = sum(luminance_values) / max(len(luminance_values), 1)
    avg_sat = sum(saturation_values) / max(len(saturation_values), 1)

    mood = "контрастные тёмные сцены" if avg_luma < 128 else "светлый редакторский холст с яркими акцентами"
    color_mode = "умеренно насыщенные corporate-цвета" if avg_sat < 90 else "более насыщенная технологичная палитра"
    full_bleed = "full-bleed визуалы на весь слайд" if media_ratio >= 0.9 else "комбинация full-bleed и модульной композиции"

    return (
        f"{FALLBACK_STYLE_BRIEF} "
        f"Локальный профиль эталонов: {decks} deck(s), среднее {avg_slides:.1f} слайдов, media/slide={media_ratio:.2f}. "
        f"Преобладает {full_bleed}; визуальное настроение: {mood}; палитра: {color_mode}. "
        "При верстке: крупный фон, минимум служебных подписей, короткие тезисы, чёткие модульные блоки."
    )


def _reference_paths() -> list[str]:
    raw = os.getenv("REFERENCE_PRESENTATION_FILES", "").strip()
    if raw:
        return [item.strip() for item in raw.split(os.pathsep) if item.strip()]
    return DEFAULT_REFERENCE_PRESENTATIONS


def _source_signature(paths: list[str]) -> list[dict]:
    signature: list[dict] = []
    for raw_path in paths:
        path = Path(raw_path)
        if not path.exists():
            continue
        stat = path.stat()
        signature.append(
            {
                "path": str(path),
                "size": stat.st_size,
                "mtime": int(stat.st_mtime),
            }
        )
    return signature


def _load_cached_brief(signature: list[dict]) -> str:
    try:
        if not CACHE_PATH.exists():
            return ""
        payload = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        if payload.get("signature") != signature:
            return ""
        return str(payload.get("brief") or "").strip()
    except Exception:
        return ""


def _save_cached_brief(signature: list[dict], brief: str) -> None:
    try:
        CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        CACHE_PATH.write_text(
            json.dumps({"signature": signature, "brief": brief}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _extract_reference_images(pptx_path: str, output_dir: str) -> list[str]:
    extracted: list[tuple[int, str]] = []
    deck_name = Path(pptx_path).stem

    with ZipFile(pptx_path) as archive:
        media_names = sorted(
            name for name in archive.namelist()
            if name.startswith("ppt/media/") and name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))
        )

        for media_name in media_names:
            try:
                data = archive.read(media_name)
                ext = Path(media_name).suffix.lower() or ".png"
                out_path = Path(output_dir) / f"{deck_name}_{Path(media_name).stem}{ext}"
                out_path.write_bytes(data)

                with Image.open(out_path) as img:
                    width, height = img.size
                    if width < 1200 or height < 600:
                        out_path.unlink(missing_ok=True)
                        continue

                    # Remove the lower-right corner where NotebookLM watermark usually lives.
                    cropped = img.crop((0, 0, int(width * 0.93), int(height * 0.95)))
                    cropped.save(out_path)
                    extracted.append((width * height, str(out_path)))
            except Exception:
                continue

    extracted.sort(key=lambda item: item[0], reverse=True)
    return [path for _, path in extracted[:1]]


def _build_reference_images() -> list[str]:
    temp_dir = tempfile.mkdtemp(prefix="ref_style_", dir="temp")
    collected: list[str] = []

    try:
        for pptx_path in _reference_paths()[:MAX_REFERENCE_DECKS]:
            if not Path(pptx_path).exists():
                continue
            collected.extend(_extract_reference_images(pptx_path, temp_dir))

        if not collected:
            shutil.rmtree(temp_dir, ignore_errors=True)
            return []

        return collected
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return []


def _cleanup_reference_images(image_paths: list[str]) -> None:
    if not image_paths:
        return
    root = Path(image_paths[0]).parent
    shutil.rmtree(root, ignore_errors=True)


def _analyze_reference_images(image_paths: list[str]) -> str:
    client = get_artemox_client()
    if client is None or not image_paths:
        return ""

    images: list[Image.Image] = []
    try:
        for path in image_paths:
            images.append(Image.open(path).convert("RGB"))

        prompt = (
            "Ты анализируешь визуальные референсы очень качественных презентаций. "
            "Игнорируй любые служебные подписи, watermark и надпись NotebookLM в правом нижнем углу. "
            "Не пересказывай тему конкретных слайдов. "
            "Нужен только style brief для генерации презентаций похожего уровня.\n\n"
            "Верни краткий структурированный текст на русском из 5 блоков:\n"
            "1. Уровень и общее впечатление.\n"
            "2. Композиция и плотность: сколько воздуха, как распределяются крупные формы, "
            "насколько короткий текст уместен.\n"
            "3. Палитра, свет, визуальные мотивы, тип графики и метафор.\n"
            "4. Типичные архетипы слайдов: hero, cards, process, chart+commentary, matrix, radial, comparison, map/cluster.\n"
            "5. Чего избегать, чтобы не скатиться ниже уровня референсов.\n\n"
            "Сфокусируйся на повторяющихся паттернах между референсами."
        )

        response = client.models.generate_content(
            model=ARTEMOX_MEDIA_MODEL,
            contents=[prompt, *images],
        )
        return str(getattr(response, "text", "") or "").strip()
    except Exception as exc:
        print(f"Reference style analysis failed: {exc}")
        return ""
    finally:
        for image in images:
            try:
                image.close()
            except Exception:
                pass


def get_reference_style_brief() -> str:
    if ULTRA_CHEAP_MODE or DISABLE_REFERENCE_MODEL_ANALYSIS:
        return _local_reference_brief(_reference_paths())

    signature = _source_signature(_reference_paths())
    if not signature:
        return FALLBACK_STYLE_BRIEF

    cached = _load_cached_brief(signature)
    if cached:
        return cached

    image_paths = _build_reference_images()
    if not image_paths:
        return FALLBACK_STYLE_BRIEF

    try:
        brief = _analyze_reference_images(image_paths) or FALLBACK_STYLE_BRIEF
        _save_cached_brief(signature, brief)
        return brief
    finally:
        _cleanup_reference_images(image_paths)
