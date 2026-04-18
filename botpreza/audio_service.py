"""Audio Overview: NotebookLM-like host+guest podcast rendered from slide deck.

Pipeline:
1. script_from_slides(slides) — builds a host/guest Russian-language dialogue
   via Gemini using speaker_notes + key content.
2. synthesize_audio(script) — renders each turn through Gemini TTS, stitches
   WAV segments into a single MP3 (or WAV fallback if ffmpeg is unavailable).

Everything is resilient: if TTS model is unavailable, the module returns None
and the caller should fall back to PPTX-only delivery.
"""

import asyncio
import io
import json
import os
import re
import struct
import subprocess
import tempfile
import wave
from typing import Any

from artemox_client import get_artemox_client

AUDIO_ENABLED = os.getenv("AUDIO_OVERVIEW_ENABLED", "1").strip().lower() in {"1", "true", "yes", "on"}
AUDIO_SCRIPT_MODEL = os.getenv("ARTEMOX_AUDIO_SCRIPT_MODEL", "").strip()
AUDIO_TTS_MODEL = os.getenv("ARTEMOX_TTS_MODEL", "gemini-2.5-flash-preview-tts").strip()
AUDIO_SCRIPT_TIMEOUT = int(os.getenv("AUDIO_SCRIPT_TIMEOUT", "90") or "90")
AUDIO_TTS_TIMEOUT = int(os.getenv("AUDIO_TTS_TIMEOUT", "60") or "60")
AUDIO_MAX_TURNS = int(os.getenv("AUDIO_MAX_TURNS", "18") or "18")
AUDIO_HOST_VOICE = os.getenv("AUDIO_HOST_VOICE", "Kore").strip() or "Kore"
AUDIO_GUEST_VOICE = os.getenv("AUDIO_GUEST_VOICE", "Puck").strip() or "Puck"


SCRIPT_PROMPT = (
    "Ты — сценарист подкаста Audio Overview в стиле NotebookLM. Тебе дана структура презентации "
    "со слайдами (title, bullets, stats, quotes, speaker_notes). Собери живой двухголосный диалог "
    "между ведущим (HOST) и аналитиком (GUEST), который 4-7 минут вслух пересказывает презентацию.\n\n"
    "Требования:\n"
    "1. Язык — русский, живая разговорная речь, без канцелярита и без чтения буллетов дословно.\n"
    "2. Это пересказ для слушателя, который не видит слайды: вводи контекст, объясняй цифры словами, "
    "комментируй, подчёркивай выводы.\n"
    "3. Реплики короткие: 1-3 предложения каждая. Диалог должен быть естественным: HOST задаёт вопросы и "
    "подталкивает, GUEST раскрывает смысл, иногда приводит цифры/цитаты из материала.\n"
    "4. Структура: короткий тёплый вход (1-2 реплики), разбор темы по слайдам (основная часть), "
    "финальные takeaways (1-2 реплики) и дружелюбное закрытие.\n"
    "5. Цифры и цитаты бери ТОЛЬКО из поля stats/quotes/source_hint в слайдах. Не выдумывай.\n"
    f"6. Всего {AUDIO_MAX_TURNS} реплик максимум.\n\n"
    "Верни строго валидный JSON без markdown по схеме:\n"
    '{\n'
    '  "title": "краткое название эпизода",\n'
    '  "turns": [ {"speaker": "HOST|GUEST", "text": "реплика"} ]\n'
    '}'
)


def _call_model(prompt: str, model: str) -> str:
    client = get_artemox_client()
    if client is None:
        return ""
    response = client.models.generate_content(model=model, contents=prompt)
    text = getattr(response, "text", "") or ""
    return text.strip()


def _clean_json(text: str) -> str:
    cleaned = (text or "").strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.removeprefix("```json").removeprefix("```").strip()
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3].strip()
    return cleaned


def _slides_to_brief(slides: list[dict]) -> str:
    brief_slides = []
    for slide in slides:
        content = slide.get("content") if isinstance(slide.get("content"), dict) else {}
        brief_slides.append(
            {
                "n": slide.get("slide_number"),
                "layout": slide.get("layout_type"),
                "title": slide.get("title") or content.get("title"),
                "subtitle": slide.get("subtitle") or content.get("subtitle"),
                "bullets": slide.get("bullets") or content.get("bullets") or [],
                "stats": slide.get("stats") or content.get("stats") or [],
                "quotes": slide.get("quotes") or content.get("quotes") or [],
                "source_hint": slide.get("source_hint") or content.get("source_hint"),
                "speaker_notes": slide.get("speaker_notes"),
            }
        )
    return json.dumps(brief_slides, ensure_ascii=False)


def _normalize_script(raw: str) -> dict | None:
    try:
        payload = json.loads(_clean_json(raw))
    except Exception:
        return None
    if not isinstance(payload, dict):
        return None
    turns_raw = payload.get("turns")
    if not isinstance(turns_raw, list) or not turns_raw:
        return None
    turns: list[dict] = []
    for item in turns_raw:
        if not isinstance(item, dict):
            continue
        speaker = str(item.get("speaker") or "").strip().upper()
        text = re.sub(r"\s+", " ", str(item.get("text") or "")).strip()
        if not text:
            continue
        if speaker not in ("HOST", "GUEST"):
            speaker = "HOST" if len(turns) % 2 == 0 else "GUEST"
        turns.append({"speaker": speaker, "text": text[:600]})
        if len(turns) >= AUDIO_MAX_TURNS:
            break
    if not turns:
        return None
    return {"title": str(payload.get("title") or "Audio overview")[:120], "turns": turns}


def build_script_from_slides_sync(slides: list[dict]) -> dict | None:
    if not slides:
        return None
    client = get_artemox_client()
    if client is None:
        return None
    # Fall back to brain model when script model not explicitly set.
    script_model = AUDIO_SCRIPT_MODEL or os.getenv("ARTEMOX_BRAIN_MODEL", "").strip() or os.getenv("ARTEMOX_MODEL", "").strip() or "gemini-2.5-pro"
    brief = _slides_to_brief(slides)
    prompt = f"{SCRIPT_PROMPT}\n\nСтруктура презентации (JSON):\n{brief}"
    try:
        raw = _call_model(prompt, script_model)
    except Exception as exc:
        print(f"Audio script generation failed: {exc}")
        return None
    if not raw:
        return None
    return _normalize_script(raw)


async def build_script_from_slides(slides: list[dict]) -> dict | None:
    try:
        return await asyncio.wait_for(
            asyncio.to_thread(build_script_from_slides_sync, slides),
            timeout=AUDIO_SCRIPT_TIMEOUT,
        )
    except asyncio.TimeoutError:
        print("Audio script generation timed out.")
        return None


def _collect_audio_payload(response) -> bytes | None:
    candidates = getattr(response, "candidates", None) or []
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        parts = getattr(content, "parts", None) or []
        for part in parts:
            inline = getattr(part, "inline_data", None)
            if inline is not None:
                data = getattr(inline, "data", None)
                if data:
                    if isinstance(data, (bytes, bytearray)):
                        return bytes(data)
                    if isinstance(data, str):
                        try:
                            import base64
                            return base64.b64decode(data)
                        except Exception:
                            continue
    return None


def _synthesize_turn_sync(text: str, voice: str) -> bytes | None:
    client = get_artemox_client()
    if client is None:
        return None
    try:
        from google.genai import types  # type: ignore
        config = types.GenerateContentConfig(
            response_modalities=["AUDIO"],
            speech_config=types.SpeechConfig(
                voice_config=types.VoiceConfig(
                    prebuilt_voice_config=types.PrebuiltVoiceConfig(voice_name=voice),
                )
            ),
        )
        response = client.models.generate_content(
            model=AUDIO_TTS_MODEL,
            contents=text,
            config=config,
        )
    except Exception as exc:
        print(f"TTS call failed ({voice}): {exc}")
        return None
    return _collect_audio_payload(response)


async def _synthesize_turn(text: str, voice: str) -> bytes | None:
    try:
        return await asyncio.wait_for(
            asyncio.to_thread(_synthesize_turn_sync, text, voice),
            timeout=AUDIO_TTS_TIMEOUT,
        )
    except asyncio.TimeoutError:
        print(f"TTS turn timed out ({voice})")
        return None


def _pcm_to_wav_bytes(pcm: bytes, sample_rate: int = 24000, channels: int = 1, sample_width: int = 2) -> bytes:
    buf = io.BytesIO()
    with wave.open(buf, "wb") as wav:
        wav.setnchannels(channels)
        wav.setsampwidth(sample_width)
        wav.setframerate(sample_rate)
        wav.writeframes(pcm)
    return buf.getvalue()


def _looks_like_wav(buf: bytes) -> bool:
    return len(buf) > 44 and buf[:4] == b"RIFF" and buf[8:12] == b"WAVE"


def _concat_wavs(segments: list[bytes]) -> bytes | None:
    if not segments:
        return None
    params = None
    frames = bytearray()
    for seg in segments:
        try:
            with wave.open(io.BytesIO(seg), "rb") as wav:
                cur_params = wav.getparams()
                if params is None:
                    params = cur_params
                elif cur_params[:3] != params[:3]:
                    return None
                frames.extend(wav.readframes(wav.getnframes()))
        except Exception as exc:
            print(f"Concat WAV failed: {exc}")
            return None
    if params is None:
        return None
    out = io.BytesIO()
    with wave.open(out, "wb") as wav:
        wav.setparams(params)
        wav.writeframes(bytes(frames))
    return out.getvalue()


def _try_convert_to_mp3(wav_bytes: bytes) -> tuple[bytes, str] | tuple[bytes, str]:
    """Return (bytes, ext). If ffmpeg is available, convert to mp3; otherwise return wav."""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as src:
            src.write(wav_bytes)
            src_path = src.name
        dst_path = src_path.replace(".wav", ".mp3")
        result = subprocess.run(
            ["ffmpeg", "-y", "-i", src_path, "-ac", "1", "-b:a", "96k", dst_path],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=60,
        )
        if result.returncode == 0 and os.path.exists(dst_path):
            with open(dst_path, "rb") as f:
                mp3 = f.read()
            try:
                os.remove(src_path)
                os.remove(dst_path)
            except OSError:
                pass
            return mp3, "mp3"
    except (FileNotFoundError, subprocess.TimeoutExpired, Exception) as exc:
        print(f"ffmpeg conversion skipped: {exc}")
    return wav_bytes, "wav"


async def build_audio_overview(slides: list[dict], output_path: str | None = None) -> str | None:
    """Generate an Audio Overview from slide structure. Returns path to audio file or None."""
    if not AUDIO_ENABLED:
        return None
    script = await build_script_from_slides(slides)
    if not script:
        return None
    turns = script.get("turns") or []
    if not turns:
        return None

    segments: list[bytes] = []
    for turn in turns:
        voice = AUDIO_HOST_VOICE if turn["speaker"] == "HOST" else AUDIO_GUEST_VOICE
        pcm_or_wav = await _synthesize_turn(turn["text"], voice)
        if not pcm_or_wav:
            print(f"Skipping turn (no audio): {turn['speaker']} · {turn['text'][:60]}")
            continue
        if _looks_like_wav(pcm_or_wav):
            segments.append(pcm_or_wav)
        else:
            segments.append(_pcm_to_wav_bytes(pcm_or_wav))

    if not segments:
        print("Audio overview: no audio segments produced")
        return None

    merged_wav = _concat_wavs(segments)
    if not merged_wav:
        print("Audio overview: failed to concatenate WAV segments")
        return None

    final_bytes, ext = _try_convert_to_mp3(merged_wav)
    if output_path is None:
        fd, output_path = tempfile.mkstemp(prefix="audio_overview_", suffix=f".{ext}")
        os.close(fd)
    else:
        base, _ = os.path.splitext(output_path)
        output_path = f"{base}.{ext}"
    with open(output_path, "wb") as f:
        f.write(final_bytes)
    print(f"Audio overview saved: {output_path} ({len(final_bytes)} bytes, {len(segments)} turns)")
    return output_path
