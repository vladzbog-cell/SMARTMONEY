import asyncio
import json
import logging
import os
import re
import tempfile

from aiogram import Bot, F, Router
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from ai_service import analyze_and_create_structure
from artemox_client import ARTEMOX_QUOTA_EXCEEDED_SENTINEL
from audio_service import AUDIO_ENABLED, build_audio_overview
from outline_service import build_outline_markdown
from content_extractor import (
    ARTEMOX_MEDIA_QUOTA_EXCEEDED_SENTINEL,
    extract_context_from_media,
    extract_text_from_docx,
    extract_text_from_pdf,
    extract_text_from_url,
    extract_text_from_youtube,
)
from presentation_builder import build_pptx_from_json


router = Router()
ARTEMOX_RETRY_ATTEMPTS = 3
ARTEMOX_RETRY_DELAY_SEC = 20

MAIN_MENU_TEXTS = {
    "🪄 Создать презентацию",
    "💼 Мой профиль / Лимиты",
    "📚 Примеры работ",
    "🆘 Поддержка",
}


class GenState(StatesGroup):
    waiting_for_style = State()


def _style_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏛 Академический / Отчёт", callback_data="style:academic")],
        [InlineKeyboardButton(text="🚀 Питч стартапа", callback_data="style:pitch")],
        [InlineKeyboardButton(text="🎨 Креативный", callback_data="style:creative")],
        [InlineKeyboardButton(text="🤖 На усмотрение ИИ", callback_data="style:auto")],
    ])


def _get_document_extension(message: Message) -> str:
    if not message.document or not message.document.file_name:
        return ""
    _, extension = os.path.splitext(message.document.file_name.lower())
    return extension


def _guess_suffix(message: Message) -> str:
    if message.document and message.document.file_name:
        _, extension = os.path.splitext(message.document.file_name)
        return extension or ".bin"
    if message.photo:
        return ".jpg"
    if message.video and message.video.file_name:
        _, extension = os.path.splitext(message.video.file_name)
        return extension or ".mp4"
    if message.video:
        return ".mp4"
    if message.voice:
        return ".ogg"
    if message.audio and message.audio.file_name:
        _, extension = os.path.splitext(message.audio.file_name)
        return extension or ".mp3"
    if message.audio:
        return ".mp3"
    return ".bin"


async def _download_to_tempfile(message: Message, file_id: str, suffix: str) -> str:
    telegram_file = await message.bot.get_file(file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_path = temp_file.name
    await message.bot.download(telegram_file, destination=temp_path)
    return temp_path


def _extract_first_url(text: str) -> str | None:
    match = re.search(r"https?://\S+", text)
    return match.group(0) if match else None


def _normalize_result(result) -> str:
    if isinstance(result, str):
        return result
    return json.dumps(result, ensure_ascii=False, indent=2)


def _extract_slides_for_audio(formatted_result: str) -> list[dict]:
    try:
        payload = json.loads(formatted_result)
    except Exception:
        return []
    if isinstance(payload, list):
        return [s for s in payload if isinstance(s, dict)]
    if isinstance(payload, dict):
        slides = payload.get("slides")
        if isinstance(slides, list):
            return [s for s in slides if isinstance(s, dict)]
    return []


def _is_retryable_artemox_error(text: str) -> bool:
    normalized = (text or "").lower()
    return any(marker in normalized for marker in ("401", "429", "auth_error", "authentication error", "too many requests", "quota", "limit"))


@router.message(F.document)
async def process_document(message: Message, bot: Bot, state: FSMContext):
    status_msg = await message.answer("📥 Начинаю работу: скачиваю файл...")
    try:
        file_id = message.document.file_id
        file_name = message.document.file_name or "document.bin"
        file_info = await bot.get_file(file_id)

        os.makedirs("temp", exist_ok=True)
        temp_path = f"temp/{file_name}"

        await bot.download_file(file_info.file_path, destination=temp_path)
        await status_msg.edit_text(f"⚙️ Файл скачан. Извлекаю текст из {file_name}...")

        text_content = ""
        source_kind = "text"
        source_label = file_name
        if file_name.lower().endswith('.pdf'):
            text_content = await extract_text_from_pdf(temp_path)
            source_kind = "pdf"
        elif file_name.lower().endswith('.docx'):
            text_content = await extract_text_from_docx(temp_path)
            source_kind = "docx"
        else:
            await status_msg.edit_text("❌ Я пока не умею читать этот формат файлов. Отправьте PDF или DOCX.")
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return

        if os.path.exists(temp_path):
            os.remove(temp_path)

        if not text_content or len(text_content.strip()) < 10:
            await status_msg.edit_text("❌ Не удалось извлечь текст. Возможно, файл пустой или состоит только из сложных картинок.")
            return

        await state.update_data(
            extracted_text=text_content,
            status_msg_id=status_msg.message_id,
            source_kind=source_kind,
            source_label=source_label,
        )
        await state.set_state(GenState.waiting_for_style)

        await status_msg.edit_text(
            "📄 Текст извлечён! Какой стиль оформления использовать?",
            reply_markup=_style_keyboard(),
        )

    except Exception as e:
        print(f"🔥 КРИТИЧЕСКАЯ ОШИБКА В ХЭНДЛЕРЕ: {e}")
        await status_msg.edit_text("❌ Произошла системная ошибка. Разработчик уже смотрит логи.")


@router.message(F.text | F.photo | F.voice | F.audio | F.video)
async def process_user_material(message: Message, state: FSMContext):
    if message.text and message.text.startswith("/"):
        return
    if message.text and message.text in MAIN_MENU_TEXTS:
        return

    initial_status = "📥 Скачиваю файл с серверов Telegram..." if message.document else "⏳ Принял материал! Распознаю и анализирую..."
    status_message = await message.answer(initial_status)
    temp_input_path = None
    stage_name = "подготовки материала"

    try:
        extracted_text = None
        source_kind = "text"
        source_label = ""

        if message.photo:
            stage_name = "скачивания изображения"
            temp_input_path = await _download_to_tempfile(message, message.photo[-1].file_id, ".jpg")
            stage_name = "распознавания изображения"
            await status_message.edit_text("⚙️ Читаю содержимое документа...")
            extracted_text = await extract_context_from_media(temp_input_path, "image")
            source_kind = "image"
            source_label = "изображение"
        elif message.video:
            stage_name = "скачивания видео"
            temp_input_path = await _download_to_tempfile(
                message,
                message.video.file_id,
                _guess_suffix(message),
            )
            stage_name = "анализа видео"
            await status_message.edit_text("⚙️ Анализирую видео и извлекаю смысловой контекст...")
            extracted_text = await extract_context_from_media(temp_input_path, "video")
            source_kind = "video"
            source_label = message.video.file_name or "видео"
        elif message.voice:
            stage_name = "скачивания голосового сообщения"
            temp_input_path = await _download_to_tempfile(message, message.voice.file_id, ".ogg")
            stage_name = "расшифровки аудио"
            await status_message.edit_text("⚙️ Читаю содержимое документа...")
            extracted_text = await extract_context_from_media(temp_input_path, "audio")
            source_kind = "voice"
            source_label = "голосовое"
        elif message.audio:
            stage_name = "скачивания аудиофайла"
            temp_input_path = await _download_to_tempfile(
                message,
                message.audio.file_id,
                _guess_suffix(message),
            )
            stage_name = "расшифровки аудио"
            await status_message.edit_text("⚙️ Читаю содержимое документа...")
            extracted_text = await extract_context_from_media(temp_input_path, "audio")
            source_kind = "audio"
            source_label = (message.audio.file_name if message.audio else "") or "аудио"
        elif message.text:
            stage_name = "извлечения текста"
            source_text = message.text.strip()
            url = _extract_first_url(source_text)

            if url and ("youtube.com" in url or "youtu.be" in url):
                extracted_text = await extract_text_from_youtube(url)
                source_kind = "youtube"
                source_label = url
            elif url and url.startswith("http"):
                extracted_text = await extract_text_from_url(url)
                source_kind = "url"
                source_label = url
            else:
                extracted_text = source_text
                source_kind = "text"
                source_label = "текст в сообщении"

        if extracted_text == ARTEMOX_MEDIA_QUOTA_EXCEEDED_SENTINEL:
            await status_message.edit_text(
                "❌ Лимит Artemox API для распознавания медиа временно исчерпан. Попробуйте позже или проверьте баланс ключа в панели Artemox."
            )
            return

        if not extracted_text:
            await status_message.edit_text(f"❌ Ошибка на этапе {stage_name}.")
            return

        await state.update_data(
            extracted_text=extracted_text,
            status_msg_id=status_message.message_id,
            source_kind=source_kind,
            source_label=source_label,
        )
        await state.set_state(GenState.waiting_for_style)

        await status_message.edit_text(
            "📄 Материал обработан! Какой стиль оформления использовать?",
            reply_markup=_style_keyboard(),
        )

    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
        logging.exception("Failed to process user material")
        await status_message.edit_text(f"❌ Ошибка на этапе {stage_name}.")
    finally:
        if temp_input_path and os.path.exists(temp_input_path):
            try:
                os.remove(temp_input_path)
            except OSError:
                logging.exception("Failed to remove temporary input file")


# ── Style callback: triggers generation after user picks a style ─────────

@router.callback_query(GenState.waiting_for_style, F.data.startswith("style:"))
async def on_style_chosen(callback: CallbackQuery, state: FSMContext):
    style = callback.data.split(":", 1)[1]  # academic / pitch / creative / auto
    data = await state.get_data()
    extracted_text = data.get("extracted_text", "")
    status_msg_id = data.get("status_msg_id")
    source_kind = data.get("source_kind", "text")
    source_label = data.get("source_label", "")
    await state.clear()
    await callback.answer()

    chat_id = callback.message.chat.id
    bot = callback.bot

    STYLE_LABELS = {
        "academic": "🏛 Академический",
        "pitch": "🚀 Питч стартапа",
        "creative": "🎨 Креативный",
        "auto": "🤖 На усмотрение ИИ",
    }
    style_label = STYLE_LABELS.get(style, style)

    try:
        await bot.edit_message_text(
            f"📥 Сканирую материал и применяю стиль «{style_label}»...",
            chat_id=chat_id, message_id=status_msg_id,
        )
    except Exception:
        status_msg = await bot.send_message(chat_id, f"📥 Сканирую материал и применяю стиль «{style_label}»...")
        status_msg_id = status_msg.message_id

    temp_pptx_path = None
    try:
        await bot.edit_message_text(
            "🧠 Выстраиваю повествование: контекст → находки → цифры → выводы...",
            chat_id=chat_id, message_id=status_msg_id,
        )

        result = ""
        last_retry_error = ""
        for attempt in range(1, ARTEMOX_RETRY_ATTEMPTS + 1):
            try:
                result = await analyze_and_create_structure(extracted_text, style)
                if result == ARTEMOX_QUOTA_EXCEEDED_SENTINEL:
                    raise RuntimeError("429 quota")
                if result:
                    break
            except Exception as analyze_exc:
                last_retry_error = str(analyze_exc)
                if attempt < ARTEMOX_RETRY_ATTEMPTS and _is_retryable_artemox_error(last_retry_error):
                    await bot.edit_message_text(
                        f"🧠 Анализ не прошёл (попытка {attempt}/{ARTEMOX_RETRY_ATTEMPTS}). Жду {ARTEMOX_RETRY_DELAY_SEC} сек и повторяю...",
                        chat_id=chat_id, message_id=status_msg_id,
                    )
                    await asyncio.sleep(ARTEMOX_RETRY_DELAY_SEC)
                    continue
                raise

            if attempt < ARTEMOX_RETRY_ATTEMPTS and not result:
                await bot.edit_message_text(
                    f"🧠 Получен пустой ответ (попытка {attempt}/{ARTEMOX_RETRY_ATTEMPTS}). Жду {ARTEMOX_RETRY_DELAY_SEC} сек и повторяю...",
                    chat_id=chat_id, message_id=status_msg_id,
                )
                await asyncio.sleep(ARTEMOX_RETRY_DELAY_SEC)

        if not result:
            if _is_retryable_artemox_error(last_retry_error):
                await bot.edit_message_text(
                    "❌ Artemox временно недоступен (401/429). Попробуйте позже или проверьте статус ключа у продавца.",
                    chat_id=chat_id, message_id=status_msg_id,
                )
                return
            await bot.edit_message_text(
                "❌ Нейросеть не смогла собрать структуру из этого текста.",
                chat_id=chat_id, message_id=status_msg_id,
            )
            return

        formatted_result = _normalize_result(result)

        try:
            outline_md = build_outline_markdown(formatted_result)
            if outline_md:
                await bot.send_message(chat_id, outline_md, parse_mode="Markdown")
        except Exception:
            logging.exception("Failed to send outline preview")

        await bot.edit_message_text(
            "🎨 Рисую фоны, раскладываю карточки, матрицы и статистику по слайдам...",
            chat_id=chat_id, message_id=status_msg_id,
        )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            temp_pptx_path = tmp.name

        file_path = None
        build_error = ""
        for attempt in range(1, ARTEMOX_RETRY_ATTEMPTS + 1):
            try:
                await bot.edit_message_text(
                    f"🛠 Собираю презентацию (попытка {attempt}/{ARTEMOX_RETRY_ATTEMPTS})...",
                    chat_id=chat_id, message_id=status_msg_id,
                )
                file_path = await build_pptx_from_json(
                    formatted_result,
                    temp_pptx_path,
                    meta={"source_kind": source_kind, "source_label": source_label},
                )
                break
            except Exception as build_exc:
                build_error = str(build_exc)
                if attempt < ARTEMOX_RETRY_ATTEMPTS and _is_retryable_artemox_error(build_error):
                    await bot.edit_message_text(
                        f"🛠 Сборка прервана из-за ответа шлюза (попытка {attempt}/{ARTEMOX_RETRY_ATTEMPTS}). Жду {ARTEMOX_RETRY_DELAY_SEC} сек...",
                        chat_id=chat_id, message_id=status_msg_id,
                    )
                    await asyncio.sleep(ARTEMOX_RETRY_DELAY_SEC)
                    continue
                raise

        if not file_path:
            raise RuntimeError(build_error or "build failed")

        await bot.edit_message_text(
            "✅ Презентация готова! Отправляю файл с заметками докладчика для каждого слайда...",
            chat_id=chat_id, message_id=status_msg_id,
        )
        await bot.send_document(chat_id, FSInputFile(file_path))
        await bot.edit_message_text(
            "✅ Презентация готова и отправлена. Откройте режим заметок, чтобы увидеть speaker notes к каждому слайду.",
            chat_id=chat_id, message_id=status_msg_id,
        )

        if AUDIO_ENABLED:
            audio_path = None
            try:
                slides_for_audio = _extract_slides_for_audio(formatted_result)
                if slides_for_audio:
                    await bot.edit_message_text(
                        "🎙 Собираю аудио-обзор в стиле NotebookLM: двухголосный диалог...",
                        chat_id=chat_id, message_id=status_msg_id,
                    )
                    audio_path = await build_audio_overview(slides_for_audio)
            except Exception as audio_exc:
                logging.exception("Audio overview failed: %s", audio_exc)
                audio_path = None

            if audio_path and os.path.exists(audio_path):
                try:
                    await bot.send_audio(
                        chat_id,
                        FSInputFile(audio_path),
                        caption="🎧 Audio overview · host + guest",
                    )
                    await bot.edit_message_text(
                        "✅ Презентация и аудио-обзор отправлены. Наслаждайтесь прослушиванием!",
                        chat_id=chat_id, message_id=status_msg_id,
                    )
                except Exception:
                    logging.exception("Failed to send audio overview")
                finally:
                    try:
                        os.remove(audio_path)
                    except OSError:
                        pass
    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА (on_style_chosen): {e}")
        logging.exception("Failed during generation after style choice")
        try:
            await bot.edit_message_text(
                "❌ Произошла системная ошибка. Разработчик уже смотрит логи.",
                chat_id=chat_id, message_id=status_msg_id,
            )
        except Exception:
            pass
    finally:
        if temp_pptx_path and os.path.exists(temp_pptx_path):
            try:
                os.remove(temp_pptx_path)
            except OSError:
                pass
