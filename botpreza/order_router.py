import os
import traceback

from aiogram import Bot, Router, F
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputMediaPhoto,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database import (
    check_and_use_revision,
    create_order,
    get_order,
    get_order_by_thread,
    update_order_status,
    update_thread_id,
)
from handlers import process_document, process_user_material

router = Router()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ADMIN_GROUP_ID = -1003932097395
SUPPORT_URL = "https://t.me/supprez_bot"
PAYMENT_PHONE = "+79067012117"

URGENCY_LABELS = {
    "urgency_fast": "⚡ Срочно (до 1 часа) — 250 руб.",
    "urgency_std": "🕐 Стандарт (до 3 часов) — 150 руб.",
}


class OrderState(StatesGroup):
    waiting_for_content_type = State()
    waiting_for_required_wishes = State()
    waiting_for_materials = State()
    waiting_for_format = State()
    waiting_for_urgency = State()
    waiting_for_receipt = State()


class RevisionState(StatesGroup):
    waiting_for_comment = State()


def main_menu_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🪄 Создать презентацию")],
            [KeyboardButton(text="💼 Мой профиль / Лимиты")],
            [KeyboardButton(text="📚 Примеры работ")],
            [KeyboardButton(text="🆘 Поддержка")],
        ],
        resize_keyboard=True,
    )


def cancel_action_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_to_main")],
    ])


def content_type_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 Есть свой план (0 руб)", callback_data="content_own")],
        [InlineKeyboardButton(text="✍️ Нужна генерация текста (+100 руб)", callback_data="content_gen")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_main")],
    ])


def required_wishes_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_content_type")],
    ])


def materials_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➡️ Готово, перейти дальше", callback_data="done_materials")],
        [InlineKeyboardButton(text="⏭️ Пропустить дополнительные материалы", callback_data="skip_materials")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_required_wishes")],
    ])


def format_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 PDF", callback_data="format_pdf")],
        [InlineKeyboardButton(text="📊 PPTX", callback_data="format_pptx")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_materials")],
    ])


def urgency_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=URGENCY_LABELS["urgency_fast"], callback_data="urgency_fast")],
        [InlineKeyboardButton(text=URGENCY_LABELS["urgency_std"], callback_data="urgency_std")],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_format")],
    ])


def receipt_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="back_to_urgency")],
    ])


def result_keyboard(order_id: int):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Принять работу", callback_data=f"accept_result_{order_id}")],
        [InlineKeyboardButton(text="🛠 Правки", callback_data=f"request_revision_{order_id}")],
        [InlineKeyboardButton(text="🆘 Поддержка", url=f"{SUPPORT_URL}?start=order_{order_id}")],
    ])


def make_more_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚀 Сделать еще презентацию", callback_data="start_order")],
        [InlineKeyboardButton(text="🆘 Поддержка", url=SUPPORT_URL)],
    ])


async def delete_message_safe(bot: Bot, chat_id: int, message_id: int | None):
    if not message_id:
        return
    try:
        await bot.delete_message(chat_id, message_id)
    except Exception:
        pass


async def clear_last_bot_message(bot: Bot, chat_id: int, state: FSMContext):
    data = await state.get_data()
    last_bot_message_id = data.get("last_bot_message_id")
    if last_bot_message_id:
        await delete_message_safe(bot, chat_id, last_bot_message_id)
        await state.update_data(last_bot_message_id=None)


async def send_and_track(message: Message, state: FSMContext, text: str, reply_markup=None, **kwargs):
    sent = await message.answer(text, reply_markup=reply_markup, **kwargs)
    await state.update_data(last_bot_message_id=sent.message_id)
    return sent


async def send_generation_prompt(message: Message, state: FSMContext):
    await state.set_state(OrderState.waiting_for_materials)
    await send_and_track(
        message,
        state,
        "Отправь мне исходники! Это может быть:\n"
        "📄 Документ (PDF, Word)\n"
        "🔗 Ссылка на статью или YouTube\n"
        "🎤 Голосовое сообщение\n"
        "📸 Фото конспекта\n"
        "Или просто напиши текст.",
        reply_markup=cancel_action_keyboard(),
    )


async def send_examples(bot: Bot, user_id: int):
    try:
        await bot.send_document(
            user_id,
            FSInputFile(os.path.join(BASE_DIR, "examples", "example_academic.pptx")),
            caption="🎓 Академическая презентация",
        )
    except Exception:
        traceback.print_exc()
    try:
        await bot.send_document(
            user_id,
            FSInputFile(os.path.join(BASE_DIR, "examples", "example_business.pptx")),
            caption="💼 Бизнес-презентация",
        )
    except Exception:
        traceback.print_exc()


# ── /start ───────────────────────────────────────────────────────────────────

@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    welcome_text = (
        "Привет! Я Preza Bot — твой личный ИИ-ассистент по презентациям.\n\n"
        "Я превращаю PDF, Word-файлы, статьи, YouTube, голосовые сообщения и фото "
        "в готовые PowerPoint-презентации за считанные секунды. "
        "Отправь любой материал, и я соберу из него PPTX без ручной возни."
    )
    photos = [
        os.path.join(BASE_DIR, "start_photos", "welcome_main.jpg"),
        os.path.join(BASE_DIR, "start_photos", "slide_academic.jpg"),
        os.path.join(BASE_DIR, "start_photos", "slide_business.jpg"),
    ]
    try:
        for path in photos:
            if not os.path.exists(path):
                print(f"❌ ФАЙЛ НЕ НАЙДЕН: {path}")

        media_group = [
            InputMediaPhoto(
                media=FSInputFile(photos[0]),
                caption=welcome_text + "\n\nНиже главное меню: можно сразу запустить генерацию или посмотреть примеры.",
            ),
            InputMediaPhoto(media=FSInputFile(photos[1])),
            InputMediaPhoto(media=FSInputFile(photos[2])),
        ]
        await message.answer_media_group(media=media_group)
        await message.answer(
            "Выбери действие в меню ниже.",
            reply_markup=main_menu_keyboard(),
        )
    except Exception:
        print("🔥 КРИТИЧЕСКАЯ ОШИБКА ОТПРАВКИ МЕДИА:")
        traceback.print_exc()
        await message.answer(welcome_text, reply_markup=main_menu_keyboard())


@router.message(Command("support"))
async def cmd_support(message: Message):
    await message.answer(
        "🆘 Если нужна помощь с генерацией, оплатой или готовым файлом, напишите в поддержку.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🆘 Поддержка", url=SUPPORT_URL)],
        ]),
    )


@router.message(F.text == "🪄 Создать презентацию")
async def start_generation_from_menu(message: Message, state: FSMContext):
    await state.clear()
    await clear_last_bot_message(message.bot, message.chat.id, state)
    await send_generation_prompt(message, state)


@router.message(F.text == "💼 Мой профиль / Лимиты")
async def show_profile_limits(message: Message):
    await message.answer(
        f"Ваш ID: {message.from_user.id}\nОсталось бесплатных генераций/правок: 3",
        reply_markup=main_menu_keyboard(),
    )


@router.message(F.text == "📚 Примеры работ")
async def show_examples_from_menu(message: Message, bot: Bot):
    await send_examples(bot, message.from_user.id)
    await message.answer(
        "📚 Вот несколько готовых работ. Когда будете готовы, нажмите «🪄 Создать презентацию».",
        reply_markup=main_menu_keyboard(),
    )


@router.message(F.text == "🆘 Поддержка")
async def support_from_menu(message: Message):
    await cmd_support(message)


@router.callback_query(F.data == "show_full_examples")
async def show_full_examples(callback: CallbackQuery, bot: Bot):
    await callback.answer()
    await send_examples(bot, callback.from_user.id)

    await callback.message.answer(
        "📚 Примеры отправлены. Для новой генерации используйте главное меню ниже.",
        reply_markup=main_menu_keyboard(),
    )


@router.callback_query(F.data == "start_order")
async def start_order(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await send_generation_prompt(callback.message, state)
    await callback.answer()


# ── Шаг 1: тип контента ──────────────────────────────────────────────────────

@router.callback_query(OrderState.waiting_for_content_type, F.data.in_({"content_own", "content_gen"}))
async def handle_content_type(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.update_data(content_type=callback.data, materials=[])
    await state.set_state(OrderState.waiting_for_required_wishes)
    await send_and_track(
        callback.message,
        state,
        "✍️ Сначала обязательно напишите основные пожелания для корректной генерации презентации: "
        "тему, примерное количество слайдов, стиль, цель или структуру.",
        reply_markup=required_wishes_keyboard(),
    )
    await callback.answer()


@router.callback_query(OrderState.waiting_for_content_type, F.data == "back_to_main")
async def back_to_main(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.clear()
    await callback.message.answer(
        "Главное меню снова доступно. Можно сразу отправить материал на генерацию.",
        reply_markup=main_menu_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data == "cancel_to_main")
async def cancel_to_main(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.clear()
    await callback.message.answer(
        "Действие отменено. Выберите следующий шаг в главном меню.",
        reply_markup=main_menu_keyboard(),
    )
    await callback.answer()


# ── Шаг 2: обязательные пожелания ────────────────────────────────────────────

@router.message(OrderState.waiting_for_required_wishes, F.text)
async def handle_required_wishes(message: Message, state: FSMContext):
    wishes = (message.text or "").strip()
    if not wishes:
        await clear_last_bot_message(message.bot, message.chat.id, state)
        await delete_message_safe(message.bot, message.chat.id, message.message_id)
        await send_and_track(
            message,
            state,
            "⚠️ Нужно обязательно указать пожелания для корректной генерации презентации. "
            "Напишите тему, количество слайдов, стиль или цель презентации.",
            reply_markup=required_wishes_keyboard(),
        )
        return

    await state.update_data(wishes=wishes)
    await state.set_state(OrderState.waiting_for_materials)
    await clear_last_bot_message(message.bot, message.chat.id, state)
    await delete_message_safe(message.bot, message.chat.id, message.message_id)
    await send_and_track(
        message,
        state,
        "📎 Теперь можете отправить дополнительные материалы: документы, фото, аудио, видео, ссылки или текст. "
        "Этот шаг можно пропустить.",
        reply_markup=materials_keyboard(),
    )


@router.message(OrderState.waiting_for_required_wishes)
async def handle_required_wishes_invalid(message: Message, state: FSMContext):
    await clear_last_bot_message(message.bot, message.chat.id, state)
    await delete_message_safe(message.bot, message.chat.id, message.message_id)
    await send_and_track(
        message,
        state,
        "⚠️ Нужно обязательно указать пожелания для корректной генерации презентации текстом.",
        reply_markup=required_wishes_keyboard(),
    )


@router.callback_query(OrderState.waiting_for_required_wishes, F.data == "back_to_content_type")
async def back_to_content_type(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_content_type)
    await send_and_track(
        callback.message,
        state,
        "📝 Выберите вариант работы с контентом:",
        reply_markup=content_type_keyboard(),
    )
    await callback.answer()


# ── Шаг 3: дополнительные материалы ──────────────────────────────────────────

@router.message(
    OrderState.waiting_for_materials,
    F.text | F.document | F.photo | F.audio | F.video | F.voice,
)
async def handle_materials(message: Message, state: FSMContext):
    await clear_last_bot_message(message.bot, message.chat.id, state)
    await state.clear()
    if message.document:
        await process_document(message, message.bot, state)
    else:
        await process_user_material(message, state)


@router.callback_query(OrderState.waiting_for_materials, F.data.in_({"done_materials", "skip_materials"}))
async def finish_materials(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_format)
    await send_and_track(
        callback.message,
        state,
        "📂 В каком формате вы хотите получить презентацию?",
        reply_markup=format_keyboard(),
    )
    await callback.answer()


@router.callback_query(OrderState.waiting_for_materials, F.data == "back_to_required_wishes")
async def back_to_required_wishes(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_required_wishes)
    await send_and_track(
        callback.message,
        state,
        "✍️ Сначала обязательно напишите основные пожелания для корректной генерации презентации: "
        "тему, примерное количество слайдов, стиль, цель или структуру.",
        reply_markup=required_wishes_keyboard(),
    )
    await callback.answer()


# ── Шаг 4: формат ────────────────────────────────────────────────────────────

@router.callback_query(OrderState.waiting_for_format, F.data.in_({"format_pdf", "format_pptx"}))
async def handle_format(callback: CallbackQuery, state: FSMContext):
    file_format = "PDF" if callback.data == "format_pdf" else "PPTX"
    await state.update_data(file_format=file_format)
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_urgency)
    await send_and_track(
        callback.message,
        state,
        f"📂 Формат: {file_format}\n\n⏱ Выберите срочность выполнения:",
        reply_markup=urgency_keyboard(),
    )
    await callback.answer()


@router.callback_query(OrderState.waiting_for_format, F.data == "back_to_materials")
async def back_to_materials(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_materials)
    await send_and_track(
        callback.message,
        state,
        "📎 Теперь можете отправить дополнительные материалы: документы, фото, аудио, видео, ссылки или текст. "
        "Этот шаг можно пропустить.",
        reply_markup=materials_keyboard(),
    )
    await callback.answer()


# ── Шаг 5: срочность ─────────────────────────────────────────────────────────

@router.callback_query(OrderState.waiting_for_urgency, F.data.in_({"urgency_fast", "urgency_std"}))
async def handle_urgency(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    content_price = 100 if data.get("content_type") == "content_gen" else 0
    urgency_price = 250 if callback.data == "urgency_fast" else 150
    total_price = content_price + urgency_price

    await state.update_data(urgency=callback.data, total_price=total_price)
    await state.set_state(OrderState.waiting_for_receipt)

    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await send_and_track(
        callback.message,
        state,
        f"💳 Итого к оплате: {total_price} руб.\n\n"
        f"Переведите сумму по номеру {PAYMENT_PHONE} и отправьте скриншот чека в этот чат.",
        reply_markup=receipt_keyboard(),
    )
    await callback.answer()


@router.callback_query(OrderState.waiting_for_urgency, F.data == "back_to_format")
async def back_to_format(callback: CallbackQuery, state: FSMContext):
    file_format = (await state.get_data()).get("file_format", "PDF")
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_format)
    await send_and_track(
        callback.message,
        state,
        f"📂 Текущий формат: {file_format}\n\nВыберите формат презентации:",
        reply_markup=format_keyboard(),
    )
    await callback.answer()


@router.callback_query(OrderState.waiting_for_receipt, F.data == "back_to_urgency")
async def back_to_urgency(callback: CallbackQuery, state: FSMContext):
    await clear_last_bot_message(callback.bot, callback.message.chat.id, state)
    await delete_message_safe(callback.bot, callback.message.chat.id, callback.message.message_id)
    await state.set_state(OrderState.waiting_for_urgency)
    await send_and_track(
        callback.message,
        state,
        "⏱ Выберите срочность выполнения:",
        reply_markup=urgency_keyboard(),
    )
    await callback.answer()


# ── Шаг 6: чек от клиента ────────────────────────────────────────────────────

@router.message(OrderState.waiting_for_receipt, F.photo)
async def handle_receipt(message: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    content_type = data.get("content_type")
    urgency = data.get("urgency")
    materials = data.get("materials", [])
    wishes = data.get("wishes", "Не указаны")
    total_price = data.get("total_price")
    file_format = data.get("file_format", "PDF")

    first_material = ""
    if materials:
        first_material = materials[0].get("file_id") or materials[0].get("content", "")

    order_id = await create_order(
        user_id=message.from_user.id,
        material=first_material,
        content_type=content_type,
        urgency=urgency,
        total_price=total_price,
        file_format=file_format,
    )

    forum_topic = await bot.create_forum_topic(
        chat_id=ADMIN_GROUP_ID,
        name=f"Заказ №{order_id} | {message.from_user.first_name}",
    )
    message_thread_id = forum_topic.message_thread_id
    await update_thread_id(order_id, message_thread_id)

    await clear_last_bot_message(message.bot, message.chat.id, state)
    await delete_message_safe(message.bot, message.chat.id, message.message_id)
    await state.clear()

    await message.answer("✅ Чек получен! Ожидайте подтверждения администратора.")

    user = message.from_user
    username = f"@{user.username}" if user.username else f"ID {user.id}"

    content_label = "Генерация текста (+100 руб)" if content_type == "content_gen" else "Свой план"
    urgency_label = URGENCY_LABELS.get(urgency, urgency)

    await bot.send_message(
        ADMIN_GROUP_ID,
        f"Новый заказ №{order_id}\n"
        f"Пользователь: {username}\n"
        f"Контент: {content_label}\n"
        f"Тариф: {urgency_label}\n"
        f"Формат: {file_format}\n"
        f"Пожелания: {wishes}\n"
        f"Сумма: {total_price} руб.",
        message_thread_id=message_thread_id,
    )

    for item in materials:
        item_type = item.get("type")
        caption = item.get("caption") or None
        if item_type == "document":
            await bot.send_document(ADMIN_GROUP_ID, item["file_id"], caption=caption, message_thread_id=message_thread_id)
        elif item_type == "photo":
            await bot.send_photo(ADMIN_GROUP_ID, item["file_id"], caption=caption, message_thread_id=message_thread_id)
        elif item_type == "audio":
            await bot.send_audio(ADMIN_GROUP_ID, item["file_id"], caption=caption, message_thread_id=message_thread_id)
        elif item_type == "video":
            await bot.send_video(ADMIN_GROUP_ID, item["file_id"], caption=caption, message_thread_id=message_thread_id)
        elif item_type == "voice":
            await bot.send_voice(ADMIN_GROUP_ID, item["file_id"], message_thread_id=message_thread_id)
        elif item_type == "text":
            await bot.send_message(
                ADMIN_GROUP_ID,
                "Материал от клиента:\n" + item.get("content", ""),
                message_thread_id=message_thread_id,
            )

    admin_keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text=f"Подтвердить (Заказ {order_id})",
            callback_data=f"admin_confirm_{order_id}",
        )],
        [InlineKeyboardButton(
            text=f"Отклонить (Заказ {order_id})",
            callback_data=f"admin_reject_{order_id}",
        )],
    ])

    await bot.send_photo(
        ADMIN_GROUP_ID,
        photo=message.photo[-1].file_id,
        caption=f"Чек оплаты к заказу №{order_id}",
        reply_markup=admin_keyboard,
        message_thread_id=message_thread_id,
    )


# ── Админ: подтвердить ────────────────────────────────────────────────────────

@router.callback_query(F.data.startswith("admin_confirm_"))
async def admin_confirm(callback: CallbackQuery, bot: Bot):
    order_id = int(callback.data.split("admin_confirm_")[1])
    await update_order_status(order_id, "in_progress")

    order = await get_order(order_id)
    if order:
        await bot.send_message(
            order["user_id"],
            "Оплата подтверждена! Ваш заказ взят в работу. Ожидайте готовую презентацию.",
        )

    await callback.message.edit_caption(
        caption=f"Чек оплаты к заказу №{order_id}\n\nОплата подтверждена",
        reply_markup=None,
    )
    await callback.answer("Подтверждено")


# ── Админ: отклонить ──────────────────────────────────────────────────────────

@router.callback_query(F.data.startswith("admin_reject_"))
async def admin_reject(callback: CallbackQuery, bot: Bot):
    order_id = int(callback.data.split("admin_reject_")[1])
    await update_order_status(order_id, "rejected")

    order = await get_order(order_id)
    if order:
        await bot.send_message(
            order["user_id"],
            "❌ Оплата не подтверждена. Заказ отменён. Если это ошибка — обратитесь в поддержку.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🆘 Поддержка по заказу", url=f"{SUPPORT_URL}?start=order_{order_id}")],
            ]),
        )

    await callback.message.edit_caption(
        caption=f"Чек оплаты к заказу №{order_id}\n\nЗаказ отклонён",
        reply_markup=None,
    )
    await callback.answer("Отклонено")


@router.message(F.chat.id == ADMIN_GROUP_ID, F.document)
async def admin_send_result(message: Message, bot: Bot):
    print(
        "📥 Получен документ в админ-группе:",
        {
            "thread_id": message.message_thread_id,
            "file_name": message.document.file_name,
            "mime_type": message.document.mime_type,
            "from_user": message.from_user.id if message.from_user else None,
            "is_reply": bool(message.reply_to_message),
        },
    )

    if not message.message_thread_id:
        await message.answer("⚠️ Документ получен вне темы. Отправьте файл внутрь темы заказа.")
        return

    order = await get_order_by_thread(message.message_thread_id)
    if not order:
        await message.answer("⚠️ Для этой темы заказ не найден.")
        return

    is_pptx_by_name = bool(message.document.file_name and message.document.file_name.lower().endswith(".pptx"))
    is_pptx_by_mime = message.document.mime_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    if not is_pptx_by_name and not is_pptx_by_mime:
        await message.answer("⚠️ Нужен именно файл PPTX. Отправьте презентацию в формате .pptx.")
        return

    await bot.send_document(
        order["user_id"],
        message.document.file_id,
        caption=(
            "🎉 Ваша презентация готова. Проверьте файл и выберите действие ниже.\n\n"
            "🛠 2 правки бесплатно, далее каждая правка стоит 50 руб."
        ),
        reply_markup=result_keyboard(order["id"]),
    )
    await update_order_status(order["id"], "completed")
    await message.answer("✅ Файл отправлен клиенту.")


@router.callback_query(F.data.startswith("accept_result_"))
async def accept_result(callback: CallbackQuery, bot: Bot):
    order_id = int(callback.data.split("accept_result_")[1])
    await update_order_status(order_id, "accepted")
    order = await get_order(order_id)
    if order and order.get("thread_id"):
        await bot.send_message(
            ADMIN_GROUP_ID,
            f"✅ Клиент принял работу по заказу №{order_id}.",
            message_thread_id=order["thread_id"],
        )
    await callback.message.edit_reply_markup(reply_markup=None)
    await callback.message.answer("✅ Готово!", reply_markup=make_more_keyboard())
    await callback.answer()


@router.callback_query(F.data.startswith("request_revision_"))
async def request_revision(callback: CallbackQuery, state: FSMContext):
    order_id = int(callback.data.split("request_revision_")[1])
    can_request = await check_and_use_revision(order_id)
    if not can_request:
        await callback.answer("⚠️ 2 бесплатные правки уже использованы. Следующая правка стоит 50 руб.", show_alert=True)
        await callback.message.answer(
            "🆘 Если хотите заказать дополнительную правку за 50 руб, напишите в поддержку.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🆘 Поддержка по заказу", url=f"{SUPPORT_URL}?start=order_{order_id}")],
            ]),
        )
        return

    await state.set_state(RevisionState.waiting_for_comment)
    await state.update_data(revision_order_id=order_id)
    await callback.message.answer(
        "✍️ Опишите, какие правки нужно внести в презентацию.\n\n"
        "🛠 Напоминаем: 2 правки бесплатно, далее каждая правка стоит 50 руб.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🆘 Поддержка по заказу", url=f"{SUPPORT_URL}?start=order_{order_id}")],
        ]),
    )
    await callback.answer()


@router.message(RevisionState.waiting_for_comment, F.text)
async def handle_revision_comment(message: Message, state: FSMContext, bot: Bot):
    revision_text = (message.text or "").strip()
    if not revision_text:
        await message.answer("✍️ Опишите правки текстом.")
        return

    data = await state.get_data()
    order_id = data.get("revision_order_id")
    order = await get_order(order_id)
    if not order or not order.get("thread_id"):
        await state.clear()
        await message.answer("⚠️ Не удалось определить заказ для отправки правок.")
        return

    await bot.send_message(
        ADMIN_GROUP_ID,
        f"🛠 Правки от клиента по заказу №{order_id}:\n{revision_text}",
        message_thread_id=order["thread_id"],
    )
    await update_order_status(order_id, "revision_requested")
    await state.clear()
    await message.answer("⏳ Ожидайте, скоро отправим вам презентацию с изменениями. Время ожидания до 20 минут.")
