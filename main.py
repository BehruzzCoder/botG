import asyncio
import logging
import sqlite3
from datetime import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
    FSInputFile,
)

from openpyxl import Workbook


# =========================
# CONFIG
# =========================
BOT_TOKEN = "8644151095:AAFt-KFefXiK-DrNAuJX3CnjL7NVupgjSRQ"
ADMIN_ID = 7428130809  # o'zingizni telegram id


# =========================
# LOGGING
# =========================
logging.basicConfig(level=logging.INFO)


# =========================
# DATABASE
# =========================
conn = sqlite3.connect("applications.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS applications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    full_name TEXT NOT NULL,
    phone TEXT NOT NULL,
    age INTEGER NOT NULL,
    experience TEXT NOT NULL,
    created_at TEXT NOT NULL
)
""")
conn.commit()


# =========================
# BOT / DISPATCHER
# =========================
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


# =========================
# STATES
# =========================
class ApplicationForm(StatesGroup):
    full_name = State()
    phone = State()
    age = State()
    experience = State()


# =========================
# KEYBOARDS
# =========================
phone_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📞 Raqam yuborish", request_contact=True)]
    ],
    resize_keyboard=True,
    one_time_keyboard=True,
)

experience_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Bor"), KeyboardButton(text="Yo‘q")]
    ],
    resize_keyboard=True,
    one_time_keyboard=True,
)


# =========================
# HELPERS
# =========================
def normalize_phone(phone: str) -> str:
    phone = phone.strip().replace(" ", "")
    if phone.startswith("998") and not phone.startswith("+998"):
        phone = "+" + phone
    return phone


def is_valid_phone(phone: str) -> bool:
    phone = normalize_phone(phone)
    if phone.startswith("+998") and len(phone) == 13:
        return phone[1:].isdigit()
    if phone.isdigit() and len(phone) >= 9:
        return True
    return False


def save_application(user_id: int, full_name: str, phone: str, age: int, experience: str):
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        INSERT INTO applications (user_id, full_name, phone, age, experience, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (user_id, full_name, phone, age, experience, created_at))
    conn.commit()

    return created_at


def get_all_applications():
    cursor.execute("""
        SELECT id, full_name, phone, age, experience, created_at
        FROM applications
        ORDER BY id DESC
    """)
    return cursor.fetchall()


def create_excel_file(filename: str = "arizalar.xlsx") -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Arizalar"

    ws.append(["ID", "Ismi", "Tel raqam", "Yoshi", "Sotuv tajribasi", "Sana"])

    rows = get_all_applications()
    for row in rows:
        ws.append(row)

    # Ustun kengliklari
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 22

    wb.save(filename)
    return filename


# =========================
# COMMANDS
# =========================
@dp.message(Command("start"))
async def start_handler(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "Assalomu alaykum.\n\n"
        "Ariza topshirish uchun savollarga javob bering.\n"
        "1-savol: Ismingizni kiriting."
    )
    await state.set_state(ApplicationForm.full_name)


@dp.message(Command("cancel"))
async def cancel_handler(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "Bekor qilindi.",
        reply_markup=ReplyKeyboardRemove()
    )


@dp.message(Command("excel"))
async def excel_handler(message: Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer("Bu buyruq faqat admin uchun.")
        return

    rows = get_all_applications()
    if not rows:
        await message.answer("Hozircha hech qanday ariza yo‘q.")
        return

    file_path = create_excel_file()
    file = FSInputFile(file_path)

    await message.answer_document(
        document=file,
        caption=f"Excel tayyor. Jami arizalar soni: {len(rows)} ta"
    )


# =========================
# FORM STEPS
# =========================
@dp.message(ApplicationForm.full_name)
async def full_name_handler(message: Message, state: FSMContext):
    full_name = (message.text or "").strip()

    if len(full_name) < 2:
        await message.answer("Ismni to‘g‘ri kiriting.")
        return

    await state.update_data(full_name=full_name)
    await message.answer(
        "2-savol: Telefon raqamingizni yuboring.",
        reply_markup=phone_keyboard
    )
    await state.set_state(ApplicationForm.phone)


@dp.message(ApplicationForm.phone, F.contact)
async def phone_contact_handler(message: Message, state: FSMContext):
    phone = message.contact.phone_number
    phone = normalize_phone(phone)

    await state.update_data(phone=phone)
    await message.answer(
        "3-savol: Yoshingiz nechida?",
        reply_markup=ReplyKeyboardRemove()
    )
    await state.set_state(ApplicationForm.age)


@dp.message(ApplicationForm.phone)
async def phone_text_handler(message: Message, state: FSMContext):
    phone = (message.text or "").strip()

    if not is_valid_phone(phone):
        await message.answer("Telefon raqamni to‘g‘ri kiriting. Masalan: +998901234567")
        return

    phone = normalize_phone(phone)
    await state.update_data(phone=phone)
    await message.answer(
        "3-savol: Yoshingiz nechida?",
        reply_markup=ReplyKeyboardRemove()
    )
    await state.set_state(ApplicationForm.age)


@dp.message(ApplicationForm.age)
async def age_handler(message: Message, state: FSMContext):
    age_text = (message.text or "").strip()

    if not age_text.isdigit():
        await message.answer("Yoshni son bilan kiriting.")
        return

    age = int(age_text)

    if age < 10 or age > 100:
        await message.answer("Yoshni to‘g‘ri kiriting.")
        return

    await state.update_data(age=age)
    await message.answer(
        "4-savol: Sotuv tajribangiz bormi?",
        reply_markup=experience_keyboard
    )
    await state.set_state(ApplicationForm.experience)


@dp.message(ApplicationForm.experience)
async def experience_handler(message: Message, state: FSMContext):
    experience = (message.text or "").strip()

    allowed = ["Bor", "Yo‘q", "Yoq", "bor", "yo‘q", "yoq"]
    if experience not in allowed:
        await message.answer("Iltimos, 'Bor' yoki 'Yo‘q' tugmasidan birini tanlang.")
        return

    if experience.lower() in ["yoq", "yo‘q"]:
        experience = "Yo‘q"
    else:
        experience = "Bor"

    data = await state.get_data()

    full_name = data["full_name"]
    phone = data["phone"]
    age = data["age"]

    created_at = save_application(
        user_id=message.from_user.id,
        full_name=full_name,
        phone=phone,
        age=age,
        experience=experience
    )

    await message.answer(
        "Arizangiz qabul qilindi ✅",
        reply_markup=ReplyKeyboardRemove()
    )

    admin_text = (
        f"📥 Yangi ariza!\n\n"
        f"👤 Ismi: {full_name}\n"
        f"📞 Tel raqam: {phone}\n"
        f"🎂 Yoshi: {age}\n"
        f"💼 Sotuv tajribasi: {experience}\n"
        f"🕒 Vaqti: {created_at}\n"
        f"🆔 User ID: {message.from_user.id}"
    )

    await bot.send_message(ADMIN_ID, admin_text)
    await state.clear()


# =========================
# MAIN
# =========================
async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())