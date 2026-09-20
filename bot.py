"""
ربات تلگرامی ثبت‌نام کلاس‌های کنکور ریاضی
--------------------------------------------
پیاده‌سازی با aiogram 3.x + SQLite
نسخه‌ی آماده‌ی استقرار روی Render.com (Web Service, پلن رایگان)

نصب پیش‌نیازها:
    pip install -r requirements.txt

اجرا (لوکال):
    python bot.py

روی Render.com، متغیرهای محیطی BOT_TOKEN و ADMIN_IDS از پنل Environment Variables
خوانده می‌شوند. اگر تنظیم نشوند، مقادیر پیش‌فرض هاردکد شده‌ی زیر (فقط برای تست) استفاده می‌شود.
"""

import asyncio
import logging
import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
    ReplyKeyboardRemove,
)

from aiohttp import web
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ============================================================
#  تنظیمات اصلی
#  اولویت با متغیرهای محیطی (Environment Variables) است؛
#  مقادیر پیش‌فرض زیر فقط برای اجرای سریع و تست لوکال هستند.
# ============================================================
BOT_TOKEN = os.environ.get("BOT_TOKEN", "PUT_YOUR_BOT_TOKEN_HERE")

_admin_ids_raw = os.environ.get("ADMIN_IDS", "111111111")
ADMIN_IDS = {int(x.strip()) for x in _admin_ids_raw.split(",") if x.strip()}

# پورتی که Render برای Health Check به آن درخواست می‌زند
PORT = int(os.environ.get("PORT", 8080))

DB_PATH = "konkur_bot.db"
EXPORT_DIR = Path("exports")
EXPORT_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("konkur_bot")

# فونت فارسی برای PDF — این فایل باید کنار bot.py و در همان ریپازیتوری باشد
FONT_PATH = Path(__file__).parent / "Vazirmatn-Regular.ttf"
PDF_FONT_NAME = "Vazirmatn"
if FONT_PATH.exists():
    pdfmetrics.registerFont(TTFont(PDF_FONT_NAME, str(FONT_PATH)))
else:
    logger.warning(
        "فایل فونت %s پیدا نشد؛ خروجی PDF بدون فونت فارسی درست کار نخواهد کرد.",
        FONT_PATH,
    )


def rtl(text: str) -> str:
    """آماده‌سازی متن فارسی/عربی برای نمایش صحیح راست‌به‌چپ در PDF."""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


router = Router()

IRAN_PHONE_REGEX = re.compile(r"^(?:0|\+98|0098)?9\d{9}$")


# ============================================================
#  لایه دیتابیس
# ============================================================
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            day TEXT NOT NULL,
            time TEXT NOT NULL,
            capacity_total INTEGER NOT NULL,
            capacity_left INTEGER NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER NOT NULL,
            telegram_id INTEGER NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            grade TEXT NOT NULL,
            major TEXT NOT NULL,
            registered_at TEXT NOT NULL,
            paid INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (class_id) REFERENCES classes (id)
        )
        """
    )
    conn.commit()

    # اگر هیچ کلاسی وجود نداشت، چند کلاس نمونه برای دمو اضافه می‌شود
    cur.execute("SELECT COUNT(*) AS c FROM classes")
    if cur.fetchone()["c"] == 0:
        sample_classes = [
            ("حسابان جامع کنکور (پایه تا کنکور)", "شنبه و دوشنبه", "۱۸:۰۰ - ۲۰:۰۰", 20),
            ("ریاضی تجربی جمع‌بندی نهایی", "یکشنبه و سه‌شنبه", "۱۷:۰۰ - ۱۹:۰۰", 15),
            ("هندسه تحلیلی و گسسته فشرده", "پنجشنبه", "۱۰:۰۰ - ۱۳:۰۰", 12),
        ]
        cur.executemany(
            """
            INSERT INTO classes (title, day, time, capacity_total, capacity_left, is_active)
            VALUES (?, ?, ?, ?, ?, 1)
            """,
            [(t, d, tm, cap, cap) for (t, d, tm, cap) in sample_classes],
        )
        conn.commit()
    conn.close()


def get_active_classes():
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM classes WHERE is_active = 1 ORDER BY id"
    ).fetchall()
    conn.close()
    return rows


def get_class(class_id: int):
    conn = get_conn()
    row = conn.execute("SELECT * FROM classes WHERE id = ?", (class_id,)).fetchone()
    conn.close()
    return row


def decrement_capacity(class_id: int) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT capacity_left FROM classes WHERE id = ?", (class_id,)
    )
    row = cur.fetchone()
    if not row or row["capacity_left"] <= 0:
        conn.close()
        return False
    cur.execute(
        "UPDATE classes SET capacity_left = capacity_left - 1 WHERE id = ?",
        (class_id,),
    )
    conn.commit()
    conn.close()
    return True


def add_capacity(class_id: int, amount: int) -> None:
    conn = get_conn()
    conn.execute(
        """
        UPDATE classes
        SET capacity_total = capacity_total + ?,
            capacity_left = capacity_left + ?
        WHERE id = ?
        """,
        (amount, amount, class_id),
    )
    conn.commit()
    conn.close()


def close_class(class_id: int) -> None:
    conn = get_conn()
    conn.execute("UPDATE classes SET is_active = 0 WHERE id = ?", (class_id,))
    conn.commit()
    conn.close()


def save_student(class_id: int, telegram_id: int, first_name: str, last_name: str,
                  phone: str, grade: str, major: str) -> int:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO students
            (class_id, telegram_id, first_name, last_name, phone, grade, major, registered_at, paid)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
        """,
        (
            class_id, telegram_id, first_name, last_name, phone, grade, major,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ),
    )
    conn.commit()
    student_id = cur.lastrowid
    conn.close()
    return student_id


def get_all_students():
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT s.*, c.title AS class_title
        FROM students s
        JOIN classes c ON c.id = s.class_id
        ORDER BY s.registered_at DESC
        """
    ).fetchall()
    conn.close()
    return rows


def get_stats():
    conn = get_conn()
    total_students = conn.execute("SELECT COUNT(*) c FROM students").fetchone()["c"]
    per_class = conn.execute(
        """
        SELECT c.title, c.capacity_total, c.capacity_left, COUNT(s.id) AS registered
        FROM classes c
        LEFT JOIN students s ON s.class_id = c.id
        WHERE c.is_active = 1
        GROUP BY c.id
        ORDER BY c.id
        """
    ).fetchall()
    conn.close()
    return total_students, per_class


# ============================================================
#  حالت‌های فرم ثبت‌نام (FSM)
# ============================================================
class RegisterForm(StatesGroup):
    full_name = State()
    phone = State()
    grade = State()
    major = State()
    confirm_payment = State()


# ============================================================
#  کیبوردهای شیشه‌ای
# ============================================================
def classes_keyboard() -> InlineKeyboardMarkup:
    classes = get_active_classes()
    buttons = []
    for c in classes:
        label = f"📘 {c['title']} | ظرفیت باقی‌مانده: {c['capacity_left']}"
        buttons.append(
            [InlineKeyboardButton(text=label, callback_data=f"class:view:{c['id']}")]
        )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def class_detail_keyboard(class_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ ثبت‌نام در این کلاس", callback_data=f"class:register:{class_id}")],
            [InlineKeyboardButton(text="🔙 بازگشت به لیست کلاس‌ها", callback_data="class:list")],
        ]
    )


def payment_keyboard(class_id: int) -> InlineKeyboardMarkup:
    fake_gateway_url = f"https://zarinpal.com/pg/StartPay/DEMO-{class_id}-SESSION"
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💳 پرداخت شهریه از درگاه", url=fake_gateway_url)],
            [InlineKeyboardButton(text="✅ پرداخت را انجام دادم", callback_data=f"pay:confirm:{class_id}")],
            [InlineKeyboardButton(text="❌ انصراف", callback_data="class:list")],
        ]
    )


def admin_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📊 آمار ثبت‌نام‌ها", callback_data="admin:stats")],
            [InlineKeyboardButton(text="📥 دریافت خروجی اکسل", callback_data="admin:export")],
            [InlineKeyboardButton(text="📄 دریافت خروجی PDF", callback_data="admin:export_pdf")],
            [InlineKeyboardButton(text="⚙️ مدیریت ظرفیت کلاس‌ها", callback_data="admin:manage")],
        ]
    )


def admin_manage_keyboard() -> InlineKeyboardMarkup:
    classes = get_active_classes()
    buttons = []
    for c in classes:
        buttons.append(
            [
                InlineKeyboardButton(text=f"➕ {c['title'][:20]}", callback_data=f"admin:addcap:{c['id']}"),
                InlineKeyboardButton(text="🔒 بستن", callback_data=f"admin:close:{c['id']}"),
            ]
        )
    buttons.append([InlineKeyboardButton(text="🔙 بازگشت", callback_data="admin:menu")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


# ============================================================
#  هندلرهای دانش‌آموز
# ============================================================
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext) -> None:
    await state.clear()
    text = (
        "🌟 <b>سلام و درود!</b>\n\n"
        "به ربات ثبت‌نام کلاس‌های تقویتی کنکور ریاضی خوش آمدید. 📐📊\n"
        "از این طریق می‌توانید در کلاس‌های فعال ثبت‌نام کرده و شهریه را به‌صورت آنلاین پرداخت کنید.\n\n"
        "👇 برای مشاهده کلاس‌های موجود روی دکمه زیر بزنید."
    )
    kb = InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="📚 مشاهده کلاس‌های فعال", callback_data="class:list")]]
    )
    await message.answer(text, reply_markup=kb)


@router.callback_query(F.data == "class:list")
async def show_classes(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    classes = get_active_classes()
    if not classes:
        await callback.message.edit_text("😔 در حال حاضر کلاس فعالی وجود ندارد.")
        await callback.answer()
        return
    await callback.message.edit_text(
        "📋 <b>لیست کلاس‌های فعال:</b>\n\nبرای مشاهده جزئیات هر کلاس روی آن بزنید 👇",
        reply_markup=classes_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("class:view:"))
async def view_class(callback: CallbackQuery) -> None:
    class_id = int(callback.data.split(":")[2])
    c = get_class(class_id)
    if not c or not c["is_active"]:
        await callback.answer("این کلاس دیگر فعال نیست.", show_alert=True)
        return

    if c["capacity_left"] <= 0:
        text = (
            f"📘 <b>{c['title']}</b>\n\n"
            f"📅 روزهای برگزاری: {c['day']}\n"
            f"⏰ ساعت برگزاری: {c['time']}\n\n"
            "🔴 متأسفانه ظرفیت این کلاس تکمیل شده است."
        )
        await callback.message.edit_text(text, reply_markup=class_detail_keyboard(class_id))
        await callback.answer()
        return

    text = (
        f"📘 <b>{c['title']}</b>\n\n"
        f"📅 روزهای برگزاری: {c['day']}\n"
        f"⏰ ساعت برگزاری: {c['time']}\n"
        f"🟢 ظرفیت باقی‌مانده: {c['capacity_left']} نفر از {c['capacity_total']}\n\n"
        "برای ثبت‌نام روی دکمه زیر بزنید ✍️"
    )
    await callback.message.edit_text(text, reply_markup=class_detail_keyboard(class_id))
    await callback.answer()


@router.callback_query(F.data.startswith("class:register:"))
async def start_registration(callback: CallbackQuery, state: FSMContext) -> None:
    class_id = int(callback.data.split(":")[2])
    c = get_class(class_id)
    if not c or not c["is_active"] or c["capacity_left"] <= 0:
        await callback.answer("ظرفیت این کلاس تکمیل شده یا کلاس غیرفعال است.", show_alert=True)
        return

    await state.update_data(class_id=class_id)
    await state.set_state(RegisterForm.full_name)
    await callback.message.edit_text(
        f"✍️ <b>ثبت‌نام در کلاس: {c['title']}</b>\n\n"
        "لطفاً <b>نام و نام خانوادگی</b> خود را وارد کنید:"
    )
    await callback.answer()


@router.message(RegisterForm.full_name)
async def process_full_name(message: Message, state: FSMContext) -> None:
    name = (message.text or "").strip()
    if len(name.split()) < 2:
        await message.answer("⚠️ لطفاً نام و نام خانوادگی را کامل و با فاصله وارد کنید. مثال: علی رضایی")
        return
    parts = name.split(maxsplit=1)
    await state.update_data(first_name=parts[0], last_name=parts[1])
    await state.set_state(RegisterForm.phone)
    await message.answer(
        "📱 لطفاً <b>شماره موبایل</b> خود را وارد کنید (مثال: ۰۹۱۲۳۴۵۶۷۸۹):",
        reply_markup=ReplyKeyboardRemove(),
    )


@router.message(RegisterForm.phone)
async def process_phone(message: Message, state: FSMContext) -> None:
    phone_raw = (message.text or "").strip()
    # تبدیل ارقام فارسی/عربی به انگلیسی
    persian_digits = "۰۱۲۳۴۵۶۷۸۹"
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    translation = str.maketrans(
        persian_digits + arabic_digits, "0123456789" + "0123456789"
    )
    phone = phone_raw.translate(translation).replace(" ", "").replace("-", "")

    if not IRAN_PHONE_REGEX.match(phone):
        await message.answer(
            "⚠️ شماره موبایل وارد شده معتبر نیست.\n"
            "لطفاً شماره را به‌صورت صحیح وارد کنید (مثال: 09123456789)."
        )
        return

    await state.update_data(phone=phone)
    await state.set_state(RegisterForm.grade)
    await message.answer("🎓 لطفاً <b>پایه تحصیلی</b> خود را وارد کنید (مثلاً: دوازدهم):")


@router.message(RegisterForm.grade)
async def process_grade(message: Message, state: FSMContext) -> None:
    grade = (message.text or "").strip()
    if not grade:
        await message.answer("⚠️ لطفاً پایه تحصیلی را وارد کنید.")
        return
    await state.update_data(grade=grade)
    await state.set_state(RegisterForm.major)
    await message.answer("📐 لطفاً <b>رشته تحصیلی</b> خود را وارد کنید (مثلاً: ریاضی فیزیک):")


@router.message(RegisterForm.major)
async def process_major(message: Message, state: FSMContext) -> None:
    major = (message.text or "").strip()
    if not major:
        await message.answer("⚠️ لطفاً رشته تحصیلی را وارد کنید.")
        return
    await state.update_data(major=major)
    data = await state.get_data()
    c = get_class(data["class_id"])

    summary = (
        "📝 <b>خلاصه اطلاعات ثبت‌نام شما:</b>\n\n"
        f"👤 نام و نام خانوادگی: {data['first_name']} {data['last_name']}\n"
        f"📱 شماره موبایل: {data['phone']}\n"
        f"🎓 پایه تحصیلی: {data['grade']}\n"
        f"📐 رشته تحصیلی: {major}\n"
        f"📘 کلاس: {c['title']}\n\n"
        "برای نهایی‌شدن ثبت‌نام، لطفاً شهریه را پرداخت کنید 💳"
    )
    await state.set_state(RegisterForm.confirm_payment)
    await message.answer(summary, reply_markup=payment_keyboard(data["class_id"]))


@router.callback_query(F.data.startswith("pay:confirm:"), RegisterForm.confirm_payment)
async def confirm_payment(callback: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    class_id = data["class_id"]
    c = get_class(class_id)

    if not c or not c["is_active"] or c["capacity_left"] <= 0:
        await callback.answer("متأسفانه ظرفیت این کلاس تکمیل شده است.", show_alert=True)
        await state.clear()
        return

    ok = decrement_capacity(class_id)
    if not ok:
        await callback.answer("ظرفیت این کلاس در همین لحظه تکمیل شد!", show_alert=True)
        await state.clear()
        return

    save_student(
        class_id=class_id,
        telegram_id=callback.from_user.id,
        first_name=data["first_name"],
        last_name=data["last_name"],
        phone=data["phone"],
        grade=data["grade"],
        major=data["major"],
    )

    await callback.message.edit_text(
        "🎉 <b>پرداخت با موفقیت تأیید شد!</b>\n\n"
        f"✅ ثبت‌نام شما در کلاس «{c['title']}» با موفقیت نهایی شد.\n"
        "📩 اطلاعات شما برای دبیر ارسال گردید.\n\n"
        "🙏 با آرزوی موفقیت در کنکور پیش رو 🌟"
    )
    await callback.answer("ثبت‌نام با موفقیت انجام شد ✅")

    for admin_id in ADMIN_IDS:
        try:
            await callback.bot.send_message(
                admin_id,
                "🔔 <b>ثبت‌نام جدید!</b>\n\n"
                f"👤 {data['first_name']} {data['last_name']}\n"
                f"📱 {data['phone']}\n"
                f"🎓 پایه: {data['grade']} | رشته: {data['major']}\n"
                f"📘 کلاس: {c['title']}",
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not notify admin %s: %s", admin_id, exc)

    await state.clear()


# ============================================================
#  هندلرهای پنل ادمین
# ============================================================
def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


@router.message(Command("admin"))
async def cmd_admin(message: Message) -> None:
    if not is_admin(message.from_user.id):
        await message.answer("⛔ شما دسترسی به پنل مدیریت را ندارید.")
        return
    await message.answer(
        "🛠 <b>پنل مدیریت دبیر</b>\n\nاز منوی زیر گزینه مورد نظر را انتخاب کنید:",
        reply_markup=admin_menu_keyboard(),
    )


@router.callback_query(F.data == "admin:menu")
async def admin_menu(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return
    await callback.message.edit_text(
        "🛠 <b>پنل مدیریت دبیر</b>\n\nاز منوی زیر گزینه مورد نظر را انتخاب کنید:",
        reply_markup=admin_menu_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data == "admin:stats")
async def admin_stats(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return

    total_students, per_class = get_stats()
    lines = [f"📊 <b>آمار کلی ثبت‌نام‌ها</b>\n", f"👥 تعداد کل ثبت‌نام‌شدگان: <b>{total_students}</b> نفر\n"]
    for row in per_class:
        lines.append(
            f"📘 {row['title']}\n"
            f"   ثبت‌نام‌شده: {row['registered']} | باقی‌مانده: {row['capacity_left']} از {row['capacity_total']}\n"
        )
    text = "\n".join(lines) if per_class else "کلاس فعالی موجود نیست."

    kb = InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="🔙 بازگشت", callback_data="admin:menu")]]
    )
    await callback.message.edit_text(text, reply_markup=kb)
    await callback.answer()


@router.callback_query(F.data == "admin:export")
async def admin_export(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return

    students = get_all_students()
    if not students:
        await callback.answer("هنوز هیچ دانش‌آموزی ثبت‌نام نکرده است.", show_alert=True)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "دانش‌آموزان"
    ws.sheet_view.rightToLeft = True

    headers = ["ردیف", "نام", "نام خانوادگی", "شماره موبایل", "پایه", "رشته", "کلاس", "تاریخ ثبت‌نام"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for idx, s in enumerate(students, start=1):
        ws.append(
            [
                idx,
                s["first_name"],
                s["last_name"],
                s["phone"],
                s["grade"],
                s["major"],
                s["class_title"],
                s["registered_at"],
            ]
        )

    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = max(12, length + 2)

    filename = EXPORT_DIR / f"students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)

    await callback.message.answer_document(
        FSInputFile(filename),
        caption=f"📥 خروجی اکسل لیست دانش‌آموزان\n👥 تعداد کل: {len(students)} نفر",
    )
    await callback.answer("فایل اکسل ارسال شد ✅")


@router.callback_query(F.data == "admin:export_pdf")
async def admin_export_pdf(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return

    students = get_all_students()
    if not students:
        await callback.answer("هنوز هیچ دانش‌آموزی ثبت‌نام نکرده است.", show_alert=True)
        return

    if not FONT_PATH.exists():
        await callback.answer(
            "فایل فونت فارسی پیدا نشد؛ ابتدا Vazirmatn-Regular.ttf را کنار bot.py قرار دهید.",
            show_alert=True,
        )
        return

    cell_style = ParagraphStyle(
        "cell", fontName=PDF_FONT_NAME, fontSize=9, leading=13, alignment=2
    )
    header_style = ParagraphStyle(
        "header", fontName=PDF_FONT_NAME, fontSize=10, leading=14,
        alignment=2, textColor=colors.white,
    )
    title_style = ParagraphStyle(
        "title", fontName=PDF_FONT_NAME, fontSize=15, leading=20, alignment=1
    )

    # ترتیب ستون‌ها برای نمایش صحیح راست‌به‌چپ برعکس چیده می‌شود
    # (ستونی که باید سمت راستِ جدول دیده شود، اول در لیست می‌آید)
    headers = ["تاریخ ثبت‌نام", "کلاس", "رشته", "پایه", "شماره موبایل", "نام خانوادگی", "نام", "ردیف"]
    table_data = [[Paragraph(rtl(h), header_style) for h in headers]]

    for idx, s in enumerate(students, start=1):
        row = [
            s["registered_at"],
            s["class_title"],
            s["major"],
            s["grade"],
            s["phone"],
            s["last_name"],
            s["first_name"],
            str(idx),
        ]
        table_data.append([Paragraph(rtl(v), cell_style) for v in row])

    filename = EXPORT_DIR / f"students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    doc = SimpleDocTemplate(
        str(filename), pagesize=A4,
        rightMargin=25, leftMargin=25, topMargin=30, bottomMargin=30,
    )
    table = Table(table_data, repeatRows=1, hAlign="CENTER")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    story = [
        Paragraph(rtl("لیست دانش‌آموزان ثبت‌نامی"), title_style),
        Spacer(1, 6),
        Paragraph(rtl(f"تعداد کل: {len(students)} نفر"), cell_style),
        Spacer(1, 12),
        table,
    ]
    doc.build(story)

    await callback.message.answer_document(
        FSInputFile(filename),
        caption=f"📄 خروجی PDF لیست دانش‌آموزان\n👥 تعداد کل: {len(students)} نفر",
    )
    await callback.answer("فایل PDF ارسال شد ✅")


@router.callback_query(F.data == "admin:manage")
async def admin_manage(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return
    await callback.message.edit_text(
        "⚙️ <b>مدیریت کلاس‌ها</b>\n\n"
        "➕ برای افزودن ۵ نفر ظرفیت و 🔒 برای بستن کلاس روی دکمه مربوطه بزنید:",
        reply_markup=admin_manage_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("admin:addcap:"))
async def admin_add_capacity(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return
    class_id = int(callback.data.split(":")[2])
    add_capacity(class_id, 5)
    c = get_class(class_id)
    await callback.answer(f"۵ نفر ظرفیت به «{c['title']}» اضافه شد ✅", show_alert=True)
    await callback.message.edit_reply_markup(reply_markup=admin_manage_keyboard())


@router.callback_query(F.data.startswith("admin:close:"))
async def admin_close_class(callback: CallbackQuery) -> None:
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ عدم دسترسی", show_alert=True)
        return
    class_id = int(callback.data.split(":")[2])
    close_class(class_id)
    await callback.answer("کلاس بسته شد 🔒", show_alert=True)
    await callback.message.edit_text(
        "⚙️ <b>مدیریت کلاس‌ها</b>\n\nکلاس با موفقیت بسته شد.",
        reply_markup=admin_manage_keyboard(),
    )


# ============================================================
#  سرور سلامتی (Health Check) برای Render.com
#  Render انتظار دارد سرویس روی $PORT به درخواست HTTP پاسخ بدهد،
#  وگرنه آن را "ناسالم" تشخیص داده و مدام ری‌استارتش می‌کند.
# ============================================================
async def handle_health(request: web.Request) -> web.Response:
    return web.Response(text="OK", status=200)


async def start_health_server() -> web.AppRunner:
    app = web.Application()
    app.router.add_get("/", handle_health)
    app.router.add_get("/health", handle_health)

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, host="0.0.0.0", port=PORT)
    await site.start()
    logger.info("Health check server listening on 0.0.0.0:%s", PORT)
    return runner


# ============================================================
#  اجرای ربات
# ============================================================
async def main() -> None:
    if BOT_TOKEN == "PUT_YOUR_BOT_TOKEN_HERE":
        raise RuntimeError(
            "BOT_TOKEN تنظیم نشده است. آن را در متغیرهای محیطی (Environment Variables) "
            "روی Render قرار دهید یا مقدار پیش‌فرض بالای فایل را جایگزین کنید."
        )

    init_db()

    bot = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    # سرور سلامتی و polling ربات هر دو باید همزمان و در یک event loop اجرا شوند
    health_runner = await start_health_server()

    logger.info("Bot is starting...")
    await bot.delete_webhook(drop_pending_updates=True)
    try:
        await dp.start_polling(bot)
    finally:
        await health_runner.cleanup()


if __name__ == "__main__":
    asyncio.run(main())
