import asyncio
import base64
import hashlib
import os
import shutil
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from dotenv import load_dotenv
from PIL import Image, ImageDraw, ImageFilter

from artemox_client import ARTEMOX_IMAGE_GENERATION_MODEL, ARTEMOX_MEDIA_MODEL, get_artemox_client


load_dotenv("keys.env")

STYLE_ANCHOR = (
    ", abstract architectural minimalism, soft urban gradients, premium financial report aesthetic, blurred textures, "
    "space on the left is reserved for text overlay, clean negative space, corporate colors, soft lighting, 16:9 ratio. "
    "DO NOT include any text or words in the image."
)

MODEL_TIMEOUT_SEC = 60
LOCAL_IMAGE_SIZE = (1920, 1080)
MIN_ACCEPTED_SOURCE_DIMENSION = 900
MAX_TEXT_PARTS_BEFORE_SUSPICIOUS = 6
MAX_TOTAL_TEXT_CHARS_BEFORE_SUSPICIOUS = 1200
ULTRA_CHEAP_MODE = os.getenv("ULTRA_CHEAP_MODE", "0").strip().lower() in {"1", "true", "yes", "on"}
MAX_IMAGE_MODEL_ATTEMPTS = int(os.getenv("MAX_IMAGE_MODEL_ATTEMPTS", "1" if ULTRA_CHEAP_MODE else "3"))
IMAGE_CACHE_DIR = os.path.join("temp", "image_cache")
CACHE_SCHEMA_VERSION = os.getenv("IMAGE_CACHE_VERSION", "2026-04-17-refpool-v1").strip()
REFERENCE_POOL_DIR = os.path.join("temp", "reference_pool")
MAX_REFERENCE_IMAGES = int(os.getenv("MAX_REFERENCE_IMAGES", "36"))
ENABLE_REFERENCE_POOL_FALLBACK = os.getenv("ENABLE_REFERENCE_POOL_FALLBACK", "0").strip().lower() in {"1", "true", "yes", "on"}
_REFERENCE_CANDIDATES: list[str] = []
IMAGE_MODELS = [
    ARTEMOX_IMAGE_GENERATION_MODEL,
    ARTEMOX_MEDIA_MODEL,
    "gemini-2.5-flash-image",
    "gemini-3.1-flash-image-preview",
    "gemini-3-pro-image-preview",
]
ALLOW_IMAGE_MODEL_FALLBACK = os.getenv("ALLOW_IMAGE_MODEL_FALLBACK", "0").strip().lower() in {"1", "true", "yes", "on"}
DEFAULT_REFERENCE_PRESENTATIONS = [
    "/opt/prezbot/references/The_Q-cumber_Revolution.pptx",
    "/opt/prezbot/references/Russian_Pharma_Market_Transformation.pptx",
    "/opt/prezbot/references/Russian_Pharmaceutical_Blueprint_(3).pptx",
    "/opt/prezbot/references/The_Fifth_Paradigm.pptx",
    "/opt/prezbot/references/Tula_Innovation_Blueprint.pptx",
]


def _build_image_prompt(prompt: str) -> str:
    normalized = (prompt or "").strip()
    return (
        "Generate a presentation illustration, not a generic wallpaper. "
        "The result must look like a premium NotebookLM-style editorial slide visual: "
        "conceptual, analytical, diagrammatic, elegant, clean, no people, no stock-photo look, no empty poster. "
        "No text response, only image: "
        + normalized
        + STYLE_ANCHOR
    )


def _validate_image(path: str) -> bool:
    try:
        with open(path, "rb") as f:
            header = f.read(16)
        if header[:4] == b"\x89PNG":
            return True
        if header[:3] == b"\xff\xd8\xff":
            return True
        if header[:4] == b"RIFF" and header[8:12] == b"WEBP":
            return True
        print(f"⚠️ Unknown image format, header: {header[:16].hex()}")
        return False
    except Exception:
        return False


def _decode_inline_data(data: bytes | str) -> bytes:
    raw = data if isinstance(data, bytes) else data.encode("utf-8")
    if raw[:5] == b"iVBOR" or raw[:4] in (b"/9j/", b"UklG"):
        return base64.b64decode(raw)
    return raw


def _normalize_to_slide_ratio(image: Image.Image) -> Image.Image:
    target_w, target_h = LOCAL_IMAGE_SIZE
    source = image.convert("RGB")
    src_w, src_h = source.size
    src_ratio = src_w / max(src_h, 1)
    target_ratio = target_w / target_h

    if abs(src_ratio - target_ratio) < 0.02:
        return source.resize((target_w, target_h))

    if src_ratio > target_ratio:
        new_h = target_h
        new_w = int(src_w * (target_h / src_h))
    else:
        new_w = target_w
        new_h = int(src_h * (target_w / src_w))

    resized = source.resize((new_w, new_h))

    # Shift crop slightly to the right to preserve focal subjects while keeping the left side calmer for text.
    left = max(0, int((new_w - target_w) * 0.62))
    top = max(0, int((new_h - target_h) * 0.5))
    left = min(left, max(0, new_w - target_w))
    top = min(top, max(0, new_h - target_h))
    return resized.crop((left, top, left + target_w, top + target_h))


def _looks_suspicious_text_response(text_parts: list[str], has_image_parts: bool) -> bool:
    if not text_parts:
        return False
    total_chars = sum(len(item) for item in text_parts)
    if not has_image_parts:
        return True
    if len(text_parts) > MAX_TEXT_PARTS_BEFORE_SUSPICIOUS:
        return True
    if total_chars > MAX_TOTAL_TEXT_CHARS_BEFORE_SUSPICIOUS:
        return True
    return False


def _save_normalized_image(image: Image.Image, output_path: str, slide_index: int) -> bool:
    width, height = image.size
    if min(width, height) < MIN_ACCEPTED_SOURCE_DIMENSION:
        print(f"⚠️ Slide {slide_index}: source image too small: {width}x{height}")
        return False
    normalized = _normalize_to_slide_ratio(image)
    normalized.save(output_path)
    return _validate_image(output_path)


def _prompt_palette(prompt: str) -> tuple[tuple[int, int, int], tuple[int, int, int], tuple[int, int, int]]:
    digest = hashlib.sha256((prompt or "slide").encode("utf-8")).digest()
    return (
        (182 + digest[0] % 34, 186 + digest[1] % 32, 194 + digest[2] % 28),
        (96 + digest[3] % 52, 122 + digest[4] % 48, 156 + digest[5] % 44),
        (208 + digest[6] % 36, 212 + digest[7] % 32, 220 + digest[8] % 28),
    )


def _infer_layout_family(prompt: str) -> str:
    text = (prompt or "").lower()
    if any(token in text for token in ("matrix", "quadrant", "2x2", "axes", "trade-off", "tradeoff")):
        return "matrix"
    if any(token in text for token in ("radial", "orbit", "ecosystem", "nucleus", "hub", "anatomy")):
        return "radial"
    if any(token in text for token in ("chart", "graph", "trend", "timeline", "forecast", "curve")):
        return "chart_focus"
    if any(token in text for token in ("comparison", "versus", "vs", "scenario", "choice", "option")):
        return "comparison"
    if any(token in text for token in ("process", "journey", "roadmap", "step", "phases", "pipeline")):
        return "process"
    if any(token in text for token in ("cards", "pillars", "themes", "blocks", "clusters")):
        return "cards"
    return "hero"


def _add_grain(image: Image.Image, amount: int = 12) -> Image.Image:
    noise = Image.effect_noise(image.size, amount).convert("L")
    noise = Image.eval(noise, lambda value: max(96, min(160, int(value))))
    grain = Image.merge("RGBA", (noise, noise, noise, Image.new("L", image.size, 18)))
    base = image.convert("RGBA")
    return Image.alpha_composite(base, grain)


def _draw_mesh(draw: ImageDraw.ImageDraw, width: int, height: int, *, color=(255, 255, 255, 28)) -> None:
    for offset in range(-height, width, 90):
        draw.line((offset, 0, offset + height, height), fill=color, width=2)
    for offset in range(0, width + height, 110):
        draw.line((offset, 0, offset - height, height), fill=(255, 255, 255, 16), width=1)


def _draw_hero_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    draw.rounded_rectangle(
        (int(width * 0.57), int(height * 0.12), int(width * 0.94), int(height * 0.86)),
        radius=42,
        outline=(255, 255, 255, 42),
        width=3,
    )
    draw.ellipse(
        (int(width * 0.56), int(height * 0.1), int(width * 0.96), int(height * 0.78)),
        fill=(*glow, 72),
    )
    for idx in range(6):
        top = int(height * (0.2 + idx * 0.1))
        draw.rounded_rectangle(
            (int(width * 0.6), top, int(width * 0.88), top + int(height * 0.055)),
            radius=20,
            outline=(255, 255, 255, 30),
            width=2,
        )
    for idx in range(5):
        x = int(width * (0.63 + idx * 0.06))
        draw.line((x, int(height * 0.18), x, int(height * 0.82)), fill=(*accent, 30), width=2)


def _draw_cards_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    positions = [
        (0.51, 0.16, 0.69, 0.38),
        (0.73, 0.16, 0.91, 0.38),
        (0.51, 0.46, 0.69, 0.68),
        (0.73, 0.46, 0.91, 0.68),
    ]
    for idx, (l, t, r, b) in enumerate(positions):
        fill = (*glow, 82 if idx % 2 == 0 else 62)
        draw.rounded_rectangle(
            (int(width * l), int(height * t), int(width * r), int(height * b)),
            radius=32,
            fill=fill,
            outline=(255, 255, 255, 34),
            width=2,
        )
        draw.line(
            (int(width * l), int(height * t) + 18, int(width * l) + 110, int(height * t) + 18),
            fill=(*accent, 220),
            width=5,
        )


def _draw_process_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    base_y = int(height * 0.6)
    for idx in range(4):
        left = int(width * (0.16 + idx * 0.18))
        draw.rounded_rectangle(
            (left, base_y, left + int(width * 0.14), base_y + int(height * 0.12)),
            radius=26,
            fill=(*glow, 80),
            outline=(255, 255, 255, 34),
            width=2,
        )
        if idx < 3:
            draw.line(
                (left + int(width * 0.14), base_y + int(height * 0.06), left + int(width * 0.18), base_y + int(height * 0.06)),
                fill=(255, 255, 255, 54),
                width=4,
            )
            arrow_x = left + int(width * 0.18)
            arrow_y = base_y + int(height * 0.06)
            draw.polygon(
                [(arrow_x, arrow_y), (arrow_x - 18, arrow_y - 12), (arrow_x - 18, arrow_y + 12)],
                fill=(*accent, 210),
            )


def _draw_chart_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    left = int(width * 0.48)
    top = int(height * 0.18)
    right = int(width * 0.93)
    bottom = int(height * 0.8)
    for idx in range(6):
        y = top + int((bottom - top) * idx / 5)
        draw.line((left, y, right, y), fill=(255, 255, 255, 22), width=1)
    for idx in range(6):
        x = left + int((right - left) * idx / 5)
        draw.line((x, top, x, bottom), fill=(255, 255, 255, 18), width=1)
    points = [
        (left + int((right - left) * 0.05), bottom - int((bottom - top) * 0.18)),
        (left + int((right - left) * 0.24), bottom - int((bottom - top) * 0.33)),
        (left + int((right - left) * 0.43), bottom - int((bottom - top) * 0.28)),
        (left + int((right - left) * 0.62), bottom - int((bottom - top) * 0.56)),
        (left + int((right - left) * 0.83), bottom - int((bottom - top) * 0.74)),
    ]
    draw.line(points, fill=(*accent, 220), width=6, joint="curve")
    for x, y in points:
        draw.ellipse((x - 10, y - 10, x + 10, y + 10), fill=(255, 255, 255, 240), outline=(*glow, 255), width=3)


def _draw_comparison_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    divider = int(width * 0.67)
    draw.line((divider, int(height * 0.16), divider, int(height * 0.84)), fill=(255, 255, 255, 38), width=3)
    for idx, x0 in enumerate((0.48, 0.72)):
        left = int(width * x0)
        draw.rounded_rectangle(
            (left, int(height * 0.26), left + int(width * 0.17), int(height * 0.72)),
            radius=30,
            fill=(*glow, 72 if idx == 0 else 58),
            outline=(255, 255, 255, 34),
            width=2,
        )
        for row in range(4):
            top = int(height * (0.33 + row * 0.09))
            draw.line((left + 20, top, left + int(width * 0.13), top), fill=(*accent, 160), width=3)


def _draw_matrix_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    left = int(width * 0.43)
    top = int(height * 0.18)
    right = int(width * 0.93)
    bottom = int(height * 0.82)
    mid_x = (left + right) // 2
    mid_y = (top + bottom) // 2
    draw.rounded_rectangle((left, top, right, bottom), radius=36, outline=(255, 255, 255, 34), width=3)
    draw.line((mid_x, top + 16, mid_x, bottom - 16), fill=(255, 255, 255, 42), width=3)
    draw.line((left + 16, mid_y, right - 16, mid_y), fill=(255, 255, 255, 42), width=3)
    nodes = [
        (left + 90, top + 90),
        (mid_x + 90, top + 110),
        (left + 110, mid_y + 85),
        (mid_x + 95, mid_y + 95),
    ]
    for idx, (x, y) in enumerate(nodes):
        size = 24 + idx * 3
        draw.ellipse((x, y, x + size, y + size), fill=(*glow, 255), outline=(*accent, 255), width=3)


def _draw_radial_visual(draw: ImageDraw.ImageDraw, width: int, height: int, accent, glow) -> None:
    cx = int(width * 0.72)
    cy = int(height * 0.52)
    for radius, alpha in ((210, 26), (155, 36), (90, 54)):
        draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius), outline=(255, 255, 255, alpha), width=3)
    draw.ellipse((cx - 42, cy - 42, cx + 42, cy + 42), fill=(*accent, 230))
    offsets = [(-220, -120), (220, -110), (-210, 135), (215, 140)]
    for dx, dy in offsets:
        node_x = cx + dx
        node_y = cy + dy
        draw.line((cx, cy, node_x, node_y), fill=(255, 255, 255, 38), width=3)
        draw.rounded_rectangle((node_x - 78, node_y - 42, node_x + 78, node_y + 42), radius=24, fill=(*glow, 72), outline=(255, 255, 255, 34), width=2)


def _preferred_models(prompt: str) -> list[str]:
    primary = (ARTEMOX_IMAGE_GENERATION_MODEL or "gemini-2.5-flash-image").strip()
    if ULTRA_CHEAP_MODE or not ALLOW_IMAGE_MODEL_FALLBACK:
        return [primary]

    ordered = [
        primary,
        "gemini-2.5-flash-image",
        ARTEMOX_MEDIA_MODEL,
        "gemini-3.1-flash-image-preview",
        "gemini-3-pro-image-preview",
    ]
    return [model for model in dict.fromkeys(ordered) if model]


def _cache_key(prompt: str) -> str:
    normalized = " ".join((prompt or "").split()).strip().lower()
    normalized = f"{CACHE_SCHEMA_VERSION}|{normalized}"
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _cache_path(prompt: str) -> str:
    return os.path.join(IMAGE_CACHE_DIR, f"{_cache_key(prompt)}.png")


def _restore_from_cache(prompt: str, output_path: str) -> bool:
    cached = _cache_path(prompt)
    if not os.path.exists(cached):
        return False
    try:
        shutil.copyfile(cached, output_path)
        return _validate_image(output_path)
    except Exception:
        return False


def _save_to_cache(prompt: str, image_path: str) -> None:
    try:
        os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)
        shutil.copyfile(image_path, _cache_path(prompt))
    except Exception:
        pass


def _reference_paths() -> list[str]:
    raw = os.getenv("REFERENCE_PRESENTATION_FILES", "").strip()
    if raw:
        return [item.strip() for item in raw.split(os.pathsep) if item.strip()]
    return DEFAULT_REFERENCE_PRESENTATIONS


def _build_reference_pool() -> list[str]:
    if not ENABLE_REFERENCE_POOL_FALLBACK:
        return []
    global _REFERENCE_CANDIDATES
    if _REFERENCE_CANDIDATES:
        return _REFERENCE_CANDIDATES

    os.makedirs(REFERENCE_POOL_DIR, exist_ok=True)
    candidates: list[str] = []
    for pptx_path in _reference_paths():
        path = Path(pptx_path)
        if not path.exists():
            continue
        try:
            with ZipFile(path) as archive:
                media_names = sorted(
                    name for name in archive.namelist()
                    if name.startswith("ppt/media/") and name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))
                )
                for media_name in media_names:
                    if len(candidates) >= MAX_REFERENCE_IMAGES:
                        break
                    try:
                        data = archive.read(media_name)
                        ext = Path(media_name).suffix.lower() or ".png"
                        out_name = f"{path.stem}_{Path(media_name).stem}{ext}"
                        out_path = os.path.join(REFERENCE_POOL_DIR, out_name)
                        if not os.path.exists(out_path):
                            with Image.open(BytesIO(data)) as image:
                                rgb = image.convert("RGB")
                                w, h = rgb.size
                                if w < 1200 or h < 650:
                                    continue
                                # Trim lower-right watermark area from NotebookLM exports.
                                cropped = rgb.crop((0, 0, int(w * 0.93), int(h * 0.95)))
                                normalized = _normalize_to_slide_ratio(cropped)
                                normalized.save(out_path)
                        if os.path.exists(out_path):
                            candidates.append(out_path)
                    except Exception:
                        continue
        except Exception:
            continue
        if len(candidates) >= MAX_REFERENCE_IMAGES:
            break

    _REFERENCE_CANDIDATES = list(dict.fromkeys(candidates))
    return _REFERENCE_CANDIDATES


def _generate_reference_background(prompt: str, slide_index: int) -> str | None:
    if not ENABLE_REFERENCE_POOL_FALLBACK:
        return None
    candidates = _build_reference_pool()
    if not candidates:
        return None
    try:
        os.makedirs("temp", exist_ok=True)
        output_path = os.path.join("temp", f"slide_{slide_index}.png")
        seed = int(hashlib.sha256(f"{prompt}|{slide_index}".encode("utf-8")).hexdigest(), 16)
        source = candidates[seed % len(candidates)]
        shutil.copyfile(source, output_path)
        if not _validate_image(output_path):
            return None
        return output_path
    except Exception:
        return None


def _generate_local_background(prompt: str, slide_index: int) -> str | None:
    try:
        os.makedirs("temp", exist_ok=True)
        output_path = os.path.join("temp", f"slide_{slide_index}.png")
        width, height = LOCAL_IMAGE_SIZE
        base, accent, glow = _prompt_palette(prompt)
        layout = _infer_layout_family(prompt)

        image = Image.new("RGB", LOCAL_IMAGE_SIZE, base)
        draw = ImageDraw.Draw(image, "RGBA")

        for y in range(height):
            mix = y / max(height - 1, 1)
            row = (
                int(base[0] * (1 - mix) + accent[0] * mix),
                int(base[1] * (1 - mix) + accent[1] * mix),
                int(base[2] * (1 - mix) + accent[2] * mix),
            )
            draw.line((0, y, width, y), fill=row)

        draw.rectangle((0, 0, int(width * 0.42), height), fill=(255, 255, 255, 42))
        _draw_mesh(draw, width, height)
        draw.ellipse((int(width * 0.46), int(height * 0.02), int(width * 0.98), int(height * 0.76)), fill=(*glow, 42))
        draw.ellipse((int(width * 0.64), int(height * 0.16), int(width * 0.98), int(height * 0.92)), fill=(255, 255, 255, 28))

        if layout == "cards":
            _draw_cards_visual(draw, width, height, accent, glow)
        elif layout == "process":
            _draw_process_visual(draw, width, height, accent, glow)
        elif layout == "chart_focus":
            _draw_chart_visual(draw, width, height, accent, glow)
        elif layout == "comparison":
            _draw_comparison_visual(draw, width, height, accent, glow)
        elif layout == "matrix":
            _draw_matrix_visual(draw, width, height, accent, glow)
        elif layout == "radial":
            _draw_radial_visual(draw, width, height, accent, glow)
        else:
            _draw_hero_visual(draw, width, height, accent, glow)

        image = image.filter(ImageFilter.GaussianBlur(radius=0.35))
        image = _add_grain(image, amount=12).convert("RGB")
        image.save(output_path)
        return output_path
    except Exception as exc:
        print(f"❌ Local fallback image generation failed on slide {slide_index}: {exc}")
        return None


def _extract_generated_image(response, output_path: str, slide_index: int) -> bool:
    candidates = getattr(response, "candidates", None) or []
    if not candidates:
        print(f"⚠️ No candidates in response for slide {slide_index}")
        return False

    parts = []
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        if content and getattr(content, "parts", None):
            parts.extend(content.parts)

    image_parts = []
    text_parts: list[str] = []
    for part in parts:
        if getattr(part, "text", None):
            text = str(part.text or "")
            text_parts.append(text)
            print(f"ℹ️ Slide {slide_index} got text part: {text[:100]}")
            continue
        if getattr(part, "inline_data", None) is not None:
            image_parts.append(part)

    if _looks_suspicious_text_response(text_parts, bool(image_parts)):
        print(f"⚠️ Slide {slide_index}: response looked too text-heavy, skipping")
        return False

    if not image_parts:
        print(f"⚠️ No image data in response for slide {slide_index}")
        return False

    last_image_part = image_parts[-1]

    try:
        image = last_image_part.as_image()
        return _save_normalized_image(image, output_path, slide_index)
    except Exception:
        inline_data = getattr(last_image_part, "inline_data", None)
        if inline_data is None or getattr(inline_data, "data", None) is None:
            print(f"⚠️ Slide {slide_index}: image part has no inline_data")
            return False

        image_bytes = _decode_inline_data(inline_data.data)

        try:
            image = Image.open(BytesIO(image_bytes))
            return _save_normalized_image(image, output_path, slide_index)
        except Exception:
            with open(output_path, "wb") as f:
                f.write(image_bytes)
            return _validate_image(output_path)


async def generate_slide_image(prompt: str, slide_index: int) -> str | None:
    if not prompt or prompt.strip().lower() in ("null", "none", ""):
        return _generate_local_background("generic presentation background", slide_index)

    os.makedirs("temp", exist_ok=True)
    os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)
    output_path = os.path.join("temp", f"slide_{slide_index}.png")
    if _restore_from_cache(prompt, output_path):
        print(f"♻️ Slide {slide_index} restored from cache")
        return output_path

    try:
        client = get_artemox_client()
        if client is None:
            fallback = _generate_reference_background(prompt, slide_index) or _generate_local_background(prompt, slide_index)
            if fallback:
                _save_to_cache(prompt, fallback)
            return fallback

        final_prompt = _build_image_prompt(prompt)
        for model_name in _preferred_models(prompt)[:max(1, MAX_IMAGE_MODEL_ATTEMPTS)]:
            try:
                response = await asyncio.wait_for(
                    asyncio.to_thread(
                        client.models.generate_content,
                        model=model_name,
                        contents=[final_prompt],
                    ),
                    timeout=MODEL_TIMEOUT_SEC,
                )
                if _extract_generated_image(response, output_path, slide_index):
                    print(f"✅ Slide {slide_index} generated via {model_name}")
                    _save_to_cache(prompt, output_path)
                    return output_path
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception as model_exc:
                print(f"❌ Model {model_name} failed on slide {slide_index}: {model_exc}")
    except Exception as e:
        print(f"❌ Image generation failed on slide {slide_index}: {e}")

    fallback = _generate_reference_background(prompt, slide_index) or _generate_local_background(prompt, slide_index)
    if fallback:
        _save_to_cache(prompt, fallback)
    return fallback
