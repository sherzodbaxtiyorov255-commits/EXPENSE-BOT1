"""
Xarajatlar/Balans Telegram boti (to'liq versiya 2)
------------------------------------------------------
Endi tugma bosmasdan ham oddiy gapda yozishingiz mumkin:
    "500000 Alidan qarz oldim"
    "13 mln maosh oldim"
    "200 000 Valiga qarz berdim"
Ovozli xabar yuborsangiz ham, bot uni matnga aylantirib xuddi shunday
qayta ishlaydi (agar tegishli dastur o'rnatilgan bo'lsa).

Oddiy xabar formatlari:
    -50000 taksi          -> UZS xarajat
    50000 nonushta         -> belgi bo'lmasa, XARAJAT deb hisoblanadi
    13 mln maosh            -> "maosh" so'zi borligi uchun avtomatik DAROMAD
    +200000 maosh           -> aniq belgi bilan ham yozish mumkin
    13000000 yoki 13 mln    -> ikkalasi ham bir xil summa deb tushuniladi

Buyruqlar:
    /start, /help     - yo'riqnoma va pastki menyu
    /balance          - balans + qarzlar + sof holat
    /history          - sana bilan birlashtirilgan tarix
    /stat             - matnli va diagrammali statistika
    /categories       - kategoriyalar bo'yicha jami xarajat
    /kurs 12700       - USD kursini belgilash (yoki 💱 Kurs tugmasi)
    /group_new, /group_join KOD, /group_leave, /group_info
    /qarz_oldim Ism summa izoh, /qarz_berdim Ism summa izoh
    /qarzlar, /qarz_yopish Ism
    /excel, /reset

O'rnatish (asosiy):
    pip install python-telegram-bot openpyxl matplotlib --break-system-packages

Qo'shimcha (ixtiyoriy funksiyalar uchun):
    Skrinshot o'qish:  pip install pytesseract Pillow --break-system-packages
                       + Tesseract-OCR dasturi (https://github.com/UB-Mannheim/tesseract/wiki)
    Ovozli xabar:      pip install SpeechRecognition pydub --break-system-packages
                       + ffmpeg dasturi (https://www.gyan.dev/ffmpeg/builds/, PATH'ga qo'shing)
    Bular o'rnatilmasa ham, bot qolgan hamma funksiyalar bilan ishlayveradi.

Ishga tushirish:
    export BOT_TOKEN="sizning_bot_tokeningiz"
    python3 expense_bot.py

Bot tokenini olish: Telegram-da @BotFather ga yozing -> /newbot
"""

import os
import re
import random
import string
import tempfile
import sqlite3
from datetime import datetime, timedelta

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pytesseract
    from PIL import Image

    OCR_AVAILABLE = True
    _tcmd = os.environ.get("TESSERACT_CMD")
    if _tcmd:
        pytesseract.pytesseract.tesseract_cmd = _tcmd
    elif os.name == "nt":
        _default_win_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(_default_win_path):
            pytesseract.pytesseract.tesseract_cmd = _default_win_path
except ImportError:
    OCR_AVAILABLE = False

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

try:
    import speech_recognition as sr
    from pydub import AudioSegment

    VOICE_AVAILABLE = True
except ImportError:
    VOICE_AVAILABLE = False

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "expenses.db")
DEFAULT_KURS = 12700.0  # taxminiy 1 USD = necha UZS

# ---------- Pastki menyu tugmalari ----------

BTN_BALANCE = "💰 Balans"
BTN_HISTORY = "📜 Tarix"
BTN_STAT = "📊 Statistika"
BTN_CATEGORIES = "🗂 Kategoriyalar"
BTN_DEBTS = "🤝 Qarzlar"
BTN_KURS = "💱 Kurs"
BTN_EXCEL = "📥 Excel"
BTN_HELP = "❓ Yordam"

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [
        [BTN_BALANCE, BTN_HISTORY],
        [BTN_STAT, BTN_CATEGORIES],
        [BTN_DEBTS, BTN_KURS],
        [BTN_EXCEL, BTN_HELP],
    ],
    resize_keyboard=True,
)

# ---------- Kategoriyalarni aniqlash ----------

CATEGORY_KEYWORDS = {
    "Oziq-ovqat": ["ovqat", "tushlik", "nonushta", "kechki", "restoran", "kafe",
                   "oshxona", "market", "magazin", "do'kon", "dokon", "un", "non",
                   "gosht", "go'sht", "sabzavot", "meva"],
    "Transport": ["taksi", "avtobus", "metro", "benzin", "yoqilg'i", "yoqilgi",
                  "mashina", "yo'l", "yol", "parkovka"],
    "Kommunal": ["svet", "elektr", "gaz", "suv", "kommunal", "ijara", "kvartira",
                 "kommunalka"],
    "Aloqa": ["mobil", "telefon", "internet", "wifi", "balans", "uzmobile",
              "beeline", "ucell"],
    "Sog'liq": ["dori", "shifokor", "klinika", "gospital", "dorixona", "vrach"],
    "Kiyim": ["kiyim", "futbolka", "poyabzal", "shim", "куртка", "kurtka"],
    "Ta'lim": ["kurs", "kitob", "maktab", "universitet", "repetitor", "darslik"],
    "O'yin-kulgi": ["kino", "konsert", "o'yin", "oyin", "bar", "klub", "disko"],
    "Daromad": ["maosh", "oylik", "bonus", "sovg'a", "sovga", "foyda", "kirim"],
}


def detect_category(note: str, is_income: bool) -> str:
    note_lower = note.lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw in note_lower:
                return category
    return "Daromad (boshqa)" if is_income else "Boshqa"


def is_income_text(note: str) -> bool:
    note_lower = note.lower()
    return any(kw in note_lower for kw in CATEGORY_KEYWORDS["Daromad"])


# ---------- Ma'lumotlar bazasi ----------

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            space_key TEXT NOT NULL,
            amount REAL NOT NULL,
            currency TEXT NOT NULL,
            category TEXT,
            note TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            created_by INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS user_group (
            user_id INTEGER PRIMARY KEY,
            group_id INTEGER NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS settings (
            space_key TEXT NOT NULL,
            key TEXT NOT NULL,
            value TEXT NOT NULL,
            PRIMARY KEY (space_key, key)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS debts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            space_key TEXT NOT NULL,
            person TEXT NOT NULL,
            direction TEXT NOT NULL,
            amount REAL NOT NULL,
            currency TEXT NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL,
            settled INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    _migrate_old_schema(conn)
    return conn


def _migrate_old_schema(conn):
    cols = {row[1] for row in conn.execute("PRAGMA table_info(transactions)").fetchall()}
    if "space_key" not in cols:
        conn.execute("ALTER TABLE transactions ADD COLUMN space_key TEXT")
        conn.execute("UPDATE transactions SET space_key='u' || user_id WHERE space_key IS NULL")
    if "category" not in cols:
        conn.execute("ALTER TABLE transactions ADD COLUMN category TEXT")
        conn.execute("UPDATE transactions SET category='Boshqa' WHERE category IS NULL")
    conn.commit()


def get_space_key(user_id: int) -> str:
    conn = get_db()
    row = conn.execute(
        "SELECT group_id FROM user_group WHERE user_id=?", (user_id,)
    ).fetchone()
    conn.close()
    if row:
        return f"g{row[0]}"
    return f"u{user_id}"


def add_transaction(user_id: int, amount: float, currency: str, category: str, note: str):
    space_key = get_space_key(user_id)
    conn = get_db()
    conn.execute(
        "INSERT INTO transactions (user_id, space_key, amount, currency, category, note, created_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (user_id, space_key, amount, currency, category, note,
         datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    conn.close()


def get_balance(space_key: str):
    conn = get_db()
    cur = conn.execute(
        "SELECT currency, SUM(amount) FROM transactions WHERE space_key=? GROUP BY currency",
        (space_key,),
    )
    rows = cur.fetchall()
    conn.close()
    balances = {"UZS": 0.0, "USD": 0.0}
    for currency, total in rows:
        balances[currency] = total or 0.0
    return balances


def get_kurs(space_key: str) -> float:
    conn = get_db()
    row = conn.execute(
        "SELECT value FROM settings WHERE space_key=? AND key='kurs'", (space_key,)
    ).fetchone()
    conn.close()
    return float(row[0]) if row else DEFAULT_KURS


def set_kurs(space_key: str, value: float):
    conn = get_db()
    conn.execute(
        "INSERT INTO settings (space_key, key, value) VALUES (?, 'kurs', ?) "
        "ON CONFLICT(space_key, key) DO UPDATE SET value=excluded.value",
        (space_key, str(value)),
    )
    conn.commit()
    conn.close()


def get_history(space_key: str, limit: int = 20):
    conn = get_db()
    cur = conn.execute(
        "SELECT amount, currency, category, note, created_at FROM transactions "
        "WHERE space_key=? ORDER BY id DESC LIMIT ?",
        (space_key, limit),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_period_stats(space_key: str, start_dt: datetime):
    conn = get_db()
    cur = conn.execute(
        "SELECT amount, currency, category FROM transactions "
        "WHERE space_key=? AND created_at>=?",
        (space_key, start_dt.isoformat(timespec="seconds")),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_period_stats_with_date(space_key: str, start_dt: datetime):
    conn = get_db()
    cur = conn.execute(
        "SELECT amount, currency, created_at FROM transactions "
        "WHERE space_key=? AND created_at>=?",
        (space_key, start_dt.isoformat(timespec="seconds")),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_combined_history(space_key: str, limit: int = 20):
    conn = get_db()
    tx_rows = conn.execute(
        "SELECT amount, currency, category, note, created_at FROM transactions "
        "WHERE space_key=? ORDER BY id DESC LIMIT ?",
        (space_key, limit),
    ).fetchall()
    debt_rows = conn.execute(
        "SELECT person, direction, amount, currency, note, created_at FROM debts "
        "WHERE space_key=? ORDER BY id DESC LIMIT ?",
        (space_key, limit),
    ).fetchall()
    conn.close()

    combined = []
    for amount, currency, category, note, created_at in tx_rows:
        combined.append({
            "kind": "tx", "created_at": created_at, "amount": amount,
            "currency": currency, "category": category, "note": note,
        })
    for person, direction, amount, currency, note, created_at in debt_rows:
        combined.append({
            "kind": "debt", "created_at": created_at, "person": person,
            "direction": direction, "amount": amount, "currency": currency, "note": note,
        })
    combined.sort(key=lambda x: x["created_at"], reverse=True)
    return combined[:limit]


def get_category_totals(space_key: str):
    conn = get_db()
    cur = conn.execute(
        "SELECT category, currency, SUM(amount) FROM transactions "
        "WHERE space_key=? AND amount<0 GROUP BY category, currency ORDER BY SUM(amount) ASC",
        (space_key,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def reset_space(space_key: str):
    conn = get_db()
    conn.execute("DELETE FROM transactions WHERE space_key=?", (space_key,))
    conn.commit()
    conn.close()


def create_group(user_id: int) -> str:
    code = "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    conn = get_db()
    cur = conn.execute(
        "INSERT INTO groups (code, created_by, created_at) VALUES (?, ?, ?)",
        (code, user_id, datetime.now().isoformat(timespec="seconds")),
    )
    group_id = cur.lastrowid
    conn.execute(
        "INSERT INTO user_group (user_id, group_id) VALUES (?, ?) "
        "ON CONFLICT(user_id) DO UPDATE SET group_id=excluded.group_id",
        (user_id, group_id),
    )
    conn.commit()
    conn.close()
    return code


def join_group(user_id: int, code: str) -> bool:
    conn = get_db()
    row = conn.execute("SELECT id FROM groups WHERE code=?", (code.upper(),)).fetchone()
    if not row:
        conn.close()
        return False
    group_id = row[0]
    conn.execute(
        "INSERT INTO user_group (user_id, group_id) VALUES (?, ?) "
        "ON CONFLICT(user_id) DO UPDATE SET group_id=excluded.group_id",
        (user_id, group_id),
    )
    conn.commit()
    conn.close()
    return True


def leave_group(user_id: int):
    conn = get_db()
    conn.execute("DELETE FROM user_group WHERE user_id=?", (user_id,))
    conn.commit()
    conn.close()


def get_group_info(user_id: int):
    conn = get_db()
    row = conn.execute(
        "SELECT g.code, g.id, (SELECT COUNT(*) FROM user_group WHERE group_id=g.id) "
        "FROM groups g JOIN user_group ug ON ug.group_id=g.id WHERE ug.user_id=?",
        (user_id,),
    ).fetchone()
    conn.close()
    return row


# ---------- Qarz (debitor/kreditor) ----------
# direction='oldim'  -> foydalanuvchi shu odamdan qarz oldi (u qarzdor)
# direction='berdim' -> foydalanuvchi shu odamga qarz berdi (u sizga qarzdor)

def add_debt(space_key: str, person: str, direction: str, amount: float, currency: str, note: str):
    conn = get_db()
    conn.execute(
        "INSERT INTO debts (space_key, person, direction, amount, currency, note, created_at, settled) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, 0)",
        (space_key, person, direction, amount, currency, note,
         datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    conn.close()


def get_debt_summary(space_key: str):
    conn = get_db()
    rows = conn.execute(
        "SELECT person, direction, currency, SUM(amount) FROM debts "
        "WHERE space_key=? AND settled=0 GROUP BY person, direction, currency",
        (space_key,),
    ).fetchall()
    conn.close()
    summary = {}
    for person, direction, currency, total in rows:
        summary.setdefault(person, {}).setdefault(currency, 0.0)
        if direction == "berdim":
            summary[person][currency] += total
        else:
            summary[person][currency] -= total
    return summary


def build_debts_summary_text(space_key: str) -> str:
    summary = get_debt_summary(space_key)
    lines = ["🤝 Qarzlar holati:"]
    has_open = False
    for person, cur_map in summary.items():
        parts = []
        for cur, net in cur_map.items():
            if abs(net) < 0.01:
                continue
            unit = "so'm" if cur == "UZS" else "$"
            if net > 0:
                parts.append(f"sizga {net:,.0f} {unit} qarzdor")
            else:
                parts.append(f"siz {abs(net):,.0f} {unit} qarzdorsiz")
        if parts:
            has_open = True
            lines.append(f"  {person}: " + ", ".join(parts))
    if not has_open:
        return "Qarzlar yo'q. 🎉"
    lines.append("\nYopish uchun: /qarz_yopish Ism")
    return "\n".join(lines)


def settle_person(space_key: str, person: str) -> int:
    conn = get_db()
    cur = conn.execute(
        "UPDATE debts SET settled=1 WHERE space_key=? AND LOWER(person)=LOWER(?) AND settled=0",
        (space_key, person),
    )
    changed = cur.rowcount
    conn.commit()
    conn.close()
    return changed


# ---------- Summani tahlil qilish (13 mln, 13 000 000, mingchisiz h.k.) ----------

def parse_amount_token(raw_amount: str, mult_word: str):
    digits = re.sub(r"[\s.,]", "", raw_amount)
    if not digits:
        return None
    try:
        value = float(digits)
    except ValueError:
        return None
    if mult_word:
        mw = mult_word.lower()
        if mw in ("mln", "million", "млн"):
            value *= 1_000_000
        elif mw in ("ming", "минг"):
            value *= 1_000
    return value


PATTERN = re.compile(
    r"^([+-]?)\s*(\d[\d\s.,]*)\s*(mln|million|млн|ming|минг)?\b\s*(\$|usd|so'?m|sum|uzs)?\s*(.*)$",
    re.IGNORECASE,
)


def parse_message(text: str):
    text = text.strip()
    match = PATTERN.match(text)
    if not match:
        return None

    sign, raw_amount, mult_word, currency_raw, note = match.groups()
    amount = parse_amount_token(raw_amount, mult_word)
    if amount is None:
        return None

    currency = "UZS" if (not currency_raw or currency_raw.lower() in ("sum", "so'm", "som", "uzs")) else "USD"
    note = note.strip() or "(izohsiz)"

    if sign == "+":
        is_income = True
    elif sign == "-":
        is_income = False
    else:
        is_income = is_income_text(note)

    amount = abs(amount) if is_income else -abs(amount)
    category = detect_category(note, is_income)
    return amount, currency, category, note


def format_balance_text(space_key: str) -> str:
    bal = get_balance(space_key)
    kurs = get_kurs(space_key)
    total_uzs = bal["UZS"] + bal["USD"] * kurs
    return (
        f"💰 Balans:\n"
        f"UZS: {bal['UZS']:,.0f} so'm\n"
        f"USD: {bal['USD']:,.2f} $\n"
        f"—\n"
        f"Umumiy (taxminan): {total_uzs:,.0f} so'm  (kurs: 1$ = {kurs:,.0f} so'm)"
    )


def build_full_status_text(space_key: str) -> str:
    kurs = get_kurs(space_key)
    bal = get_balance(space_key)
    balance_uzs = bal["UZS"] + bal["USD"] * kurs

    debt_summary = get_debt_summary(space_key)
    receivable_uzs = 0.0
    payable_uzs = 0.0
    for person, cur_map in debt_summary.items():
        for cur, net in cur_map.items():
            val_uzs = net * (kurs if cur == "USD" else 1)
            if val_uzs > 0:
                receivable_uzs += val_uzs
            else:
                payable_uzs += -val_uzs

    net_worth = balance_uzs + receivable_uzs - payable_uzs

    lines = [format_balance_text(space_key), ""]
    lines.append(build_debts_summary_text(space_key))
    lines.append("")
    lines.append(
        f"📌 Sof holat (balans + sizga qarzdorlar − siz qarzdorsiz):\n"
        f"{net_worth:,.0f} so'm (taxminan)"
    )
    return "\n".join(lines)


# ---------- Tabiiy tildagi qarz yozuvini aniqlash (tugmasiz) ----------
# Masalan: "500000 Alidan qarz oldim" yoki "Valiga 200000 qarz berdim"

FREEFORM_AMOUNT_RE = re.compile(r"(\d[\d\s.,]*\d|\d)\s*(mln|million|млн|ming|минг)?\b", re.IGNORECASE)
PERSON_FROM_RE = re.compile(r"\b([A-Za-zА-Яа-яЎўҚқҒғҲҳ']+?)dan\b", re.IGNORECASE)
PERSON_TO_RE = re.compile(r"\b([A-Za-zА-Яа-яЎўҚқҒғҲҳ']+?)ga\b", re.IGNORECASE)


def try_parse_freeform_debt(text: str):
    if "qarz" not in text.lower():
        return None
    text_lower = text.lower()
    if re.search(r"qarz\s*old", text_lower):
        direction = "oldim"
    elif re.search(r"qarz\s*ber", text_lower):
        direction = "berdim"
    else:
        return None

    m = FREEFORM_AMOUNT_RE.search(text)
    if not m:
        return None
    amount = parse_amount_token(m.group(1), m.group(2))
    if not amount:
        return None

    currency = "USD" if re.search(r"\$|\busd\b", text_lower) else "UZS"

    person = None
    pm = PERSON_FROM_RE.search(text) if direction == "oldim" else PERSON_TO_RE.search(text)
    if pm:
        candidate = pm.group(1)
        if candidate.lower() not in ("qarz", "shu", "u", "meni", "sen"):
            person = candidate

    note = text.strip()
    return person, direction, amount, currency, note


# ---------- Skrinshotdan summani o'qish (OCR) ----------

def extract_amount_candidates(text: str):
    candidates = []

    for m in re.finditer(r"\$\s?(\d{1,5}(?:[.,]\d{1,2})?)", text):
        candidates.append((float(m.group(1).replace(",", ".")), "USD"))
    for m in re.finditer(r"(\d{1,5}(?:[.,]\d{1,2})?)\s?\$", text):
        candidates.append((float(m.group(1).replace(",", ".")), "USD"))

    for m in re.finditer(r"\d{1,3}(?:[ ,.]\d{3}){1,4}", text):
        digits = re.sub(r"[ ,.]", "", m.group(0))
        if digits.isdigit():
            candidates.append((float(digits), "UZS"))
    for m in re.finditer(r"(?<!\d)(\d{4,9})(?!\d)", text):
        candidates.append((float(m.group(1)), "UZS"))

    seen = set()
    unique = []
    for amt, cur in candidates:
        key = (round(amt, 2), cur)
        if key not in seen:
            seen.add(key)
            unique.append((amt, cur))

    unique.sort(key=lambda x: -x[0])
    return unique[:4]


# ---------- Bot buyruqlari ----------

HELP_TEXT = (
    "Salom! Men xarajat/balans botiman.\n\n"
    "Endi tugma bosmasdan ham oddiy gapda yozishingiz mumkin:\n"
    "  -50000 taksi\n"
    "  13 mln maosh oldim\n"
    "  500000 Alidan qarz oldim\n"
    "  200 000 Valiga qarz berdim\n\n"
    "Ovozli xabar yuborsangiz ham tushunaman (agar o'rnatilgan bo'lsa).\n"
    "To'lov skrinshotini izoh bilan yuborsangiz, summani o'zim o'qiyman.\n\n"
    "Pastdagi tugmalar orqali tez foydalaning."
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT, reply_markup=MAIN_KEYBOARD)


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT, reply_markup=MAIN_KEYBOARD)


async def balance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    await update.message.reply_text(build_full_status_text(space_key))


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    entries = get_combined_history(space_key)
    if not entries:
        await update.message.reply_text("Hali yozuvlar yo'q.")
        return
    lines = ["📜 Oxirgi yozuvlar (xarajat/daromad va qarzlar):"]
    for e in entries:
        try:
            date_str = datetime.fromisoformat(e["created_at"]).strftime("%d.%m.%Y")
        except ValueError:
            date_str = e["created_at"].split("T")[0]
        if e["kind"] == "tx":
            sign = "+" if e["amount"] >= 0 else ""
            unit = "so'm" if e["currency"] == "UZS" else "$"
            lines.append(f"{date_str}  {sign}{e['amount']:,.0f} {unit}  [{e['category']}]  — {e['note']}")
        else:
            unit = "so'm" if e["currency"] == "UZS" else "$"
            if e["direction"] == "oldim":
                verb = f"{e['person']} dan {e['amount']:,.0f} {unit} qarz oldingiz"
            else:
                verb = f"{e['person']} ga {e['amount']:,.0f} {unit} qarz berdingiz"
            lines.append(f"{date_str}  🤝 {verb}  — {e['note']}")
    await update.message.reply_text("\n".join(lines))


def build_stat_chart(space_key: str) -> str:
    kurs = get_kurs(space_key)
    days = 7
    now = datetime.now()
    start = (now - timedelta(days=days - 1)).replace(hour=0, minute=0, second=0, microsecond=0)

    rows = get_period_stats_with_date(space_key, start)
    date_list = [(start + timedelta(days=i)).date() for i in range(days)]
    daily_income = {d: 0.0 for d in date_list}
    daily_expense = {d: 0.0 for d in date_list}

    for amount, currency, created_at in rows:
        d = datetime.fromisoformat(created_at).date()
        if d not in daily_income:
            continue
        val_uzs = amount * (kurs if currency == "USD" else 1)
        if val_uzs >= 0:
            daily_income[d] += val_uzs
        else:
            daily_expense[d] += -val_uzs

    labels = [d.strftime("%d.%m") for d in date_list]
    income_vals = [daily_income[d] for d in date_list]
    expense_vals = [daily_expense[d] for d in date_list]

    debt_summary = get_debt_summary(space_key)
    receivable = 0.0
    payable = 0.0
    for person, cur_map in debt_summary.items():
        for cur, net in cur_map.items():
            val = net * (kurs if cur == "USD" else 1)
            if val > 0:
                receivable += val
            else:
                payable += -val

    cat_rows = get_category_totals(space_key)
    cat_totals_uzs = {}
    for category, currency, total in cat_rows:
        val = abs(total) * (kurs if currency == "USD" else 1)
        cat_totals_uzs[category] = cat_totals_uzs.get(category, 0) + val
    top_cats = sorted(cat_totals_uzs.items(), key=lambda x: -x[1])[:5]

    fig, axes = plt.subplots(1, 3, figsize=(15, 4.8))
    ax1, ax2, ax3 = axes

    x = list(range(len(labels)))
    width = 0.35
    bars_income = ax1.bar([i - width / 2 for i in x], income_vals, width, label="Daromad", color="#2e7d32")
    bars_expense = ax1.bar([i + width / 2 for i in x], expense_vals, width, label="Xarajat", color="#c62828")
    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, rotation=45)
    ax1.set_title("So'nggi 7 kun (so'mda)")
    ax1.legend()
    ax1.grid(axis="y", alpha=0.3)
    ax1.bar_label(bars_income, fmt="%.0f", fontsize=7, rotation=90, padding=2)
    ax1.bar_label(bars_expense, fmt="%.0f", fontsize=7, rotation=90, padding=2)

    bars_debt = ax2.bar(["Sizga\nqarzdor", "Siz\nqarzdor"], [receivable, payable],
                         color=["#2e7d32", "#c62828"])
    ax2.set_title("Qarz holati (so'mda)")
    ax2.grid(axis="y", alpha=0.3)
    ax2.bar_label(bars_debt, fmt="%.0f", fontsize=9, padding=3)

    if top_cats:
        cat_labels = [c for c, v in top_cats]
        cat_values = [v for c, v in top_cats]
        colors = ["#c62828", "#ef6c00", "#f9a825", "#6a1b9a", "#0277bd"][:len(cat_labels)]
        wedges, texts, autotexts = ax3.pie(
            cat_values, labels=cat_labels, autopct=lambda p: f"{p:.0f}%\n({p/100*sum(cat_values):,.0f})",
            colors=colors, textprops={"fontsize": 7},
        )
        ax3.set_title("Xarajat kategoriyalari")
    else:
        ax3.axis("off")
        ax3.text(0.5, 0.5, "Xarajat yo'q", ha="center", va="center")

    fig.suptitle("Moliyaviy holat")
    fig.tight_layout()

    out_dir = tempfile.gettempdir()
    filepath = os.path.join(out_dir, f"stat_{space_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
    fig.savefig(filepath, dpi=120)
    plt.close(fig)
    return filepath


async def stat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    now = datetime.now()

    periods = [
        ("Bugun", now.replace(hour=0, minute=0, second=0, microsecond=0)),
        ("Shu hafta", (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)),
        ("Shu oy", now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)),
    ]

    lines = []
    for label, start_dt in periods:
        rows = get_period_stats(space_key, start_dt)
        income = {"UZS": 0.0, "USD": 0.0}
        expense = {"UZS": 0.0, "USD": 0.0}
        cat_totals = {}
        for amount, currency, category in rows:
            if amount >= 0:
                income[currency] += amount
            else:
                expense[currency] += -amount
                if currency == "UZS":
                    cat_totals[category] = cat_totals.get(category, 0) + (-amount)

        lines.append(f"📅 {label}:")
        lines.append(f"  Daromad: {income['UZS']:,.0f} so'm, {income['USD']:,.2f} $")
        lines.append(f"  Xarajat: {expense['UZS']:,.0f} so'm, {expense['USD']:,.2f} $")
        if cat_totals:
            top = sorted(cat_totals.items(), key=lambda x: -x[1])[:3]
            top_str = ", ".join(f"{c}: {v:,.0f}" for c, v in top if v > 0)
            if top_str:
                lines.append(f"  Ko'p sarflangan: {top_str}")
        lines.append("")

    await update.message.reply_text("\n".join(lines))

    if MATPLOTLIB_AVAILABLE:
        try:
            path = build_stat_chart(space_key)
            with open(path, "rb") as f:
                await update.message.reply_photo(
                    photo=f, caption="📊 Daromad/xarajat, qarz holati va kategoriyalar diagrammasi"
                )
            os.remove(path)
        except Exception:
            pass
    else:
        await update.message.reply_text(
            "📊 Diagramma ko'rish uchun: python -m pip install matplotlib --break-system-packages"
        )


async def categories(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    rows = get_category_totals(space_key)
    if not rows:
        await update.message.reply_text("Hali xarajatlar yo'q.")
        return
    lines = ["📊 Kategoriyalar bo'yicha jami xarajatlar:"]
    for category, currency, total in rows:
        unit = "so'm" if currency == "UZS" else "$"
        lines.append(f"  {category}: {abs(total):,.0f} {unit}")
    await update.message.reply_text("\n".join(lines))


async def kurs_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    cur_kurs = get_kurs(space_key)
    context.user_data["pending_kurs_input"] = True
    await update.message.reply_text(
        f"Joriy kurs: 1$ = {cur_kurs:,.0f} so'm\n\n"
        f"Yangi kursni raqam bilan yozing, masalan: 12700"
    )


async def kurs_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    if not context.args:
        await kurs_menu(update, context)
        return
    try:
        value = float(context.args[0].replace(",", ""))
        if value <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("Noto'g'ri qiymat. Masalan: /kurs 12700")
        return
    set_kurs(space_key, value)
    await update.message.reply_text(f"Kurs yangilandi: 1$ = {value:,.0f} so'm ✅")


async def group_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    code = create_group(user_id)
    await update.message.reply_text(
        f"👨‍👩‍👧 Jamoaviy balans yaratildi!\n\n"
        f"Taklif kodi: {code}\n\n"
        f"Oila a'zolaringiz shu kodni /group_join {code} bilan kiritsa, "
        f"hammangiz bitta umumiy balansni yuritasiz."
    )


async def group_join(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text("Kodni kiriting. Masalan: /group_join AB12CD")
        return
    code = context.args[0]
    ok = join_group(update.effective_user.id, code)
    if ok:
        await update.message.reply_text(
            f"✅ Jamoaviy balansga qo'shildingiz!\n\n"
            f"{format_balance_text(get_space_key(update.effective_user.id))}"
        )
    else:
        await update.message.reply_text("Bunday kod topilmadi. Qaytadan tekshirib ko'ring.")


async def group_leave(update: Update, context: ContextTypes.DEFAULT_TYPE):
    leave_group(update.effective_user.id)
    await update.message.reply_text("Jamoaviy balansdan chiqdingiz. Endi shaxsiy balansingiz ishlaydi. 👤")


async def group_info(update: Update, context: ContextTypes.DEFAULT_TYPE):
    info = get_group_info(update.effective_user.id)
    if not info:
        await update.message.reply_text("Siz hozir shaxsiy balansdasiz (jamoaga qo'shilmagansiz).")
        return
    code, group_id, member_count = info
    await update.message.reply_text(
        f"👨‍👩‍👧 Jamoa ma'lumoti:\nKod: {code}\nA'zolar soni: {member_count}"
    )


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    reset_space(space_key)
    await update.message.reply_text("Balans va tarix tozalandi. 🧹")


# --- Qarz buyruqlari (tugma va matnli) ---

async def debts_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    buttons = [
        [InlineKeyboardButton("➕ Men qarz oldim", callback_data="debtdir:oldim")],
        [InlineKeyboardButton("➕ Men qarz berdim", callback_data="debtdir:berdim")],
        [InlineKeyboardButton("📋 Ro'yxat", callback_data="debtlist")],
    ]
    await update.message.reply_text(
        "Qarzlar bo'limi — tugma bosing, YOKI to'g'ridan-to'g'ri shu tarzda yozing:\n"
        "\"500000 Alidan qarz oldim\" yoki \"200000 Valiga qarz berdim\"",
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def debt_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "debtlist":
        space_key = get_space_key(query.from_user.id)
        await query.edit_message_text(build_debts_summary_text(space_key))
        return
    direction = data.split(":")[1]
    context.user_data["pending_debt_direction"] = direction
    if direction == "oldim":
        prompt = "Kimdan va qancha qarz oldingiz? Masalan: Ali 500000 taksi uchun"
    else:
        prompt = "Kimga va qancha qarz berdingiz? Masalan: Vali 200000"
    await query.edit_message_text(prompt)


DEBT_PATTERN = re.compile(
    r"^(\S+)\s+(\d[\d\s.,]*)\s*(mln|million|млн|ming|минг)?\b\s*(\$|usd|so'?m|sum|uzs)?\s*(.*)$",
    re.IGNORECASE,
)


def parse_debt_args(text: str):
    text = text.strip()
    match = DEBT_PATTERN.match(text)
    if not match:
        return None
    person, raw_amount, mult_word, currency_raw, note = match.groups()
    amount = parse_amount_token(raw_amount, mult_word)
    if amount is None:
        return None
    currency = "UZS" if (not currency_raw or currency_raw.lower() in ("sum", "so'm", "som", "uzs")) else "USD"
    note = note.strip() or "(izohsiz)"
    return person, amount, currency, note


async def qarz_oldim_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = " ".join(context.args)
    parsed = parse_debt_args(text) if text else None
    if not parsed:
        await update.message.reply_text("Masalan: /qarz_oldim Ali 500000 taksi uchun")
        return
    person, amount, currency, note = parsed
    add_debt(get_space_key(update.effective_user.id), person, "oldim", amount, currency, note)
    unit = "so'm" if currency == "UZS" else "$"
    await update.message.reply_text(f"Qayd etildi: {person} dan {amount:,.0f} {unit} qarz oldingiz.")


async def qarz_berdim_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = " ".join(context.args)
    parsed = parse_debt_args(text) if text else None
    if not parsed:
        await update.message.reply_text("Masalan: /qarz_berdim Vali 200000")
        return
    person, amount, currency, note = parsed
    add_debt(get_space_key(update.effective_user.id), person, "berdim", amount, currency, note)
    unit = "so'm" if currency == "UZS" else "$"
    await update.message.reply_text(f"Qayd etildi: {person} ga {amount:,.0f} {unit} qarz berdingiz.")


async def debts_list_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    await update.message.reply_text(build_debts_summary_text(space_key))


async def qarz_yopish_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text("Ism kiriting. Masalan: /qarz_yopish Ali")
        return
    person = " ".join(context.args)
    space_key = get_space_key(update.effective_user.id)
    changed = settle_person(space_key, person)
    if changed:
        await update.message.reply_text(f"{person} bilan hisob-kitob yopildi. ✅")
    else:
        await update.message.reply_text(f"{person} bilan ochiq qarz topilmadi.")


# --- Excel ---

def build_excel_report(space_key: str) -> str:
    conn = get_db()
    rows = conn.execute(
        "SELECT amount, currency, category, note, created_at FROM transactions "
        "WHERE space_key=? ORDER BY id ASC",
        (space_key,),
    ).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Xarajatlar"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    normal_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", bold=True, size=10)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    green_font = Font(name="Arial", size=10, color="1E7B34")
    red_font = Font(name="Arial", size=10, color="C00000")

    headers = ["Sana", "Turi", "Kategoriya", "Summa", "Valyuta", "Izoh"]
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    r = 2
    for amount, currency, category, note, created_at in rows:
        date_str = created_at.split("T")[0]
        kind = "Daromad" if amount >= 0 else "Xarajat"
        font = green_font if amount >= 0 else red_font
        ws.cell(row=r, column=1, value=date_str).font = normal_font
        ws.cell(row=r, column=2, value=kind).font = font
        ws.cell(row=r, column=3, value=category or "").font = normal_font
        amount_cell = ws.cell(row=r, column=4, value=amount)
        amount_cell.font = font
        amount_cell.number_format = "#,##0.00"
        ws.cell(row=r, column=5, value=currency).font = normal_font
        ws.cell(row=r, column=6, value=note).font = normal_font
        for col in range(1, 7):
            ws.cell(row=r, column=col).border = border
        r += 1

    last_data_row = r - 1
    r += 1
    if last_data_row >= 2:
        for currency, offset in (("UZS", 0), ("USD", 1)):
            row_num = r + offset
            ws.cell(row=row_num, column=3, value=f"Jami ({currency})").font = bold_font
            formula = f'=SUMIFS(D2:D{last_data_row},E2:E{last_data_row},"{currency}")'
            total_cell = ws.cell(row=row_num, column=4, value=formula)
            total_cell.font = bold_font
            total_cell.number_format = "#,##0.00"
            for col in range(1, 7):
                ws.cell(row=row_num, column=col).border = border

    widths = {1: 12, 2: 10, 3: 14, 4: 16, 5: 9, 6: 28}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"

    out_dir = tempfile.gettempdir()
    filename = f"balans_{space_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = os.path.join(out_dir, filename)
    wb.save(filepath)
    return filepath


async def excel_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    space_key = get_space_key(update.effective_user.id)
    rows = get_history(space_key, limit=1)
    if not rows:
        await update.message.reply_text("Hali yozuvlar yo'q, Excel yaratib bo'lmaydi.")
        return

    await update.message.reply_text("📊 Excel fayl tayyorlanmoqda...")
    filepath = build_excel_report(space_key)
    with open(filepath, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=os.path.basename(filepath),
            caption="Balansingiz bo'yicha to'liq hisobot 📈",
        )
    os.remove(filepath)


# --- Skrinshot (rasm) qabul qilish ---

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not OCR_AVAILABLE:
        await update.message.reply_text(
            "Rasm o'qish funksiyasi hozircha o'rnatilmagan. Yordamdagi yo'riqnomaga qarang, "
            "yoki summani qo'lda yozing."
        )
        return

    photo = update.message.photo[-1]
    file = await photo.get_file()
    tmp_path = os.path.join(
        tempfile.gettempdir(), f"ocr_{update.effective_user.id}_{datetime.now().strftime('%H%M%S%f')}.jpg"
    )
    await file.download_to_drive(tmp_path)

    try:
        img = Image.open(tmp_path)
        text = pytesseract.image_to_string(img)
    except Exception:
        await update.message.reply_text(
            "Rasmni o'qib bo'lmadi. Tesseract-OCR to'g'ri o'rnatilganini tekshiring, "
            "yoki summani qo'lda yozing."
        )
        return
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    candidates = extract_amount_candidates(text)
    caption = (update.message.caption or "").strip()

    if not candidates:
        await update.message.reply_text(
            "Rasmda summani aniqlay olmadim 😕 Iltimos summani qo'lda yozing, masalan: -50000 taksi"
        )
        return

    if len(candidates) == 1 and caption:
        amt, cur = candidates[0]
        category = detect_category(caption, False)
        add_transaction(update.effective_user.id, -abs(amt), cur, category, caption)
        space_key = get_space_key(update.effective_user.id)
        unit = "so'm" if cur == "UZS" else "$"
        await update.message.reply_text(
            f"Rasmdan aniqlandi: {amt:,.0f} {unit}  [{category}]  ({caption})\n\n"
            f"{format_balance_text(space_key)}"
        )
        return

    buttons = []
    for amt, cur in candidates:
        unit = "so'm" if cur == "UZS" else "$"
        buttons.append([InlineKeyboardButton(f"{amt:,.0f} {unit}", callback_data=f"amtpick:{amt}:{cur}")])
    buttons.append([InlineKeyboardButton("✏️ Qo'lda kiritish", callback_data="amtmanual")])
    context.user_data["pending_photo_caption"] = caption
    await update.message.reply_text(
        "Rasmda bir nechta summa topildi, qaysi biri to'g'ri?",
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def photo_amount_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id

    if data == "amtmanual":
        context.user_data.pop("pending_photo_caption", None)
        await query.edit_message_text("Yaxshi, summani qo'lda yozing, masalan: -50000 taksi")
        return

    _, amt_str, cur = data.split(":")
    amt = float(amt_str)
    caption = context.user_data.pop("pending_photo_caption", "") or ""

    if caption:
        category = detect_category(caption, False)
        add_transaction(user_id, -abs(amt), cur, category, caption)
        space_key = get_space_key(user_id)
        unit = "so'm" if cur == "UZS" else "$"
        await query.edit_message_text(
            f"Qabul qilindi: {amt:,.0f} {unit}  [{category}]  ({caption})\n\n"
            f"{format_balance_text(space_key)}"
        )
    else:
        context.user_data["pending_amount"] = (amt, cur)
        unit = "so'm" if cur == "UZS" else "$"
        await query.edit_message_text(
            f"Summasi: {amt:,.0f} {unit}\nEndi nima uchun ekanini yozing (masalan: taksi)"
        )


# --- Markaziy matn qayta ishlash (matn xabar va ovozli xabar shu yerga tushadi) ---

async def process_free_text(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    text = text.strip()
    user_id = update.effective_user.id
    space_key = get_space_key(user_id)

    # 1) Kurs kiritilishi kutilayotgan bo'lsa
    if context.user_data.get("pending_kurs_input"):
        try:
            value = float(re.sub(r"[\s,]", "", text))
            if value <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("Noto'g'ri raqam. Masalan: 12700")
            return
        set_kurs(space_key, value)
        context.user_data.pop("pending_kurs_input", None)
        await update.message.reply_text(f"Kurs yangilandi: 1$ = {value:,.0f} so'm ✅")
        return

    # 2) Tugma orqali boshlangan qarz kiritish jarayoni
    pending_dir = context.user_data.get("pending_debt_direction")
    if pending_dir:
        parsed = parse_debt_args(text)
        if not parsed:
            await update.message.reply_text("Tushunmadim. Masalan: Ali 500000 taksi uchun")
            return
        person, amount, currency, note = parsed
        add_debt(space_key, person, pending_dir, amount, currency, note)
        context.user_data.pop("pending_debt_direction", None)
        unit = "so'm" if currency == "UZS" else "$"
        verb = "siz qarzdorsiz" if pending_dir == "oldim" else "sizga qarzdor"
        await update.message.reply_text(f"Qayd etildi: {person} — {amount:,.0f} {unit} ({verb}) ✅")
        return

    # 3) Tabiiy tilda qarz aniqlandi, lekin ism topilmadi -> ism kutilmoqda
    pending_person_info = context.user_data.get("pending_debt_awaiting_person")
    if pending_person_info:
        direction, amount, currency, note = pending_person_info
        person = text.strip()
        if not person:
            await update.message.reply_text("Ismni yozing, masalan: Ali")
            return
        add_debt(space_key, person, direction, amount, currency, note)
        context.user_data.pop("pending_debt_awaiting_person", None)
        unit = "so'm" if currency == "UZS" else "$"
        verb = "siz qarzdorsiz" if direction == "oldim" else "sizga qarzdor"
        await update.message.reply_text(f"Qayd etildi: {person} — {amount:,.0f} {unit} ({verb}) ✅")
        return

    # 4) Rasmdan summa tanlangandan keyin izoh kutilayotgan bo'lsa
    pending_amount = context.user_data.get("pending_amount")
    if pending_amount:
        amt, cur = pending_amount
        note = text or "(izohsiz)"
        category = detect_category(note, False)
        add_transaction(user_id, -abs(amt), cur, category, note)
        context.user_data.pop("pending_amount", None)
        unit = "so'm" if cur == "UZS" else "$"
        await update.message.reply_text(
            f"Qabul qilindi: {amt:,.0f} {unit}  [{category}]  ({note})\n\n"
            f"{format_balance_text(space_key)}"
        )
        return

    # 5) Tabiiy tilda yozilgan qarz xabari ("500000 Alidan qarz oldim")
    freeform_debt = try_parse_freeform_debt(text)
    if freeform_debt:
        person, direction, amount, currency, note = freeform_debt
        if person:
            add_debt(space_key, person, direction, amount, currency, note)
            unit = "so'm" if currency == "UZS" else "$"
            verb = "siz qarzdorsiz" if direction == "oldim" else "sizga qarzdor"
            await update.message.reply_text(f"Qayd etildi: {person} — {amount:,.0f} {unit} ({verb}) ✅")
        else:
            context.user_data["pending_debt_awaiting_person"] = (direction, amount, currency, note)
            question = "Kimdan qarz oldingiz? Ismini yozing." if direction == "oldim" else "Kimga qarz berdingiz? Ismini yozing."
            await update.message.reply_text(question)
        return

    # 6) Oddiy xarajat/daromad yozuvi
    parsed = parse_message(text)
    if not parsed:
        await update.message.reply_text(
            "Tushunmadim 🤔 Masalan shunday yozing: -50000 taksi, 13 mln maosh, yoki 500000 Alidan qarz oldim"
        )
        return

    amount, currency, category, note = parsed
    add_transaction(user_id, amount, currency, category, note)

    sign = "+" if amount >= 0 else ""
    unit = "so'm" if currency == "UZS" else "$"
    await update.message.reply_text(
        f"Qabul qilindi: {sign}{amount:,.0f} {unit}  [{category}]  ({note})\n\n"
        f"{format_balance_text(space_key)}"
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    button_routes = {
        BTN_BALANCE: balance,
        BTN_HISTORY: history,
        BTN_STAT: stat,
        BTN_CATEGORIES: categories,
        BTN_DEBTS: debts_menu,
        BTN_KURS: kurs_menu,
        BTN_EXCEL: excel_export,
        BTN_HELP: help_cmd,
    }
    if text in button_routes:
        await button_routes[text](update, context)
        return

    await process_free_text(update, context, text)


# --- Ovozli xabar ---

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not VOICE_AVAILABLE:
        await update.message.reply_text(
            "Ovozli xabarni tushunish funksiyasi hozircha o'rnatilmagan. "
            "Yordamdagi yo'riqnomaga qarang, yoki matn bilan yozing."
        )
        return

    voice = update.message.voice
    file = await voice.get_file()
    base = os.path.join(tempfile.gettempdir(), f"voice_{update.effective_user.id}_{datetime.now().strftime('%H%M%S%f')}")
    ogg_path = base + ".ogg"
    wav_path = base + ".wav"
    await file.download_to_drive(ogg_path)

    text = None
    try:
        audio = AudioSegment.from_file(ogg_path)
        audio.export(wav_path, format="wav")
        recognizer = sr.Recognizer()
        with sr.AudioFile(wav_path) as source:
            audio_data = recognizer.record(source)
        try:
            text = recognizer.recognize_google(audio_data, language="uz-UZ")
        except Exception:
            try:
                text = recognizer.recognize_google(audio_data, language="ru-RU")
            except Exception:
                text = None
    except Exception:
        text = None
    finally:
        for p in (ogg_path, wav_path):
            if os.path.exists(p):
                os.remove(p)

    if not text:
        await update.message.reply_text(
            "Ovozli xabarni tanib bo'lmadi 😕 FFmpeg o'rnatilganini tekshiring, yoki matn bilan yozing."
        )
        return

    await update.message.reply_text(f"🎙 Eshitdim: \"{text}\"")
    await process_free_text(update, context, text)


def main():
    token = os.environ.get("BOT_TOKEN")
    if not token:
        raise SystemExit("BOT_TOKEN muhit o'zgaruvchisi topilmadi. export BOT_TOKEN=... qiling.")

    app = ApplicationBuilder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("balance", balance))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("stat", stat))
    app.add_handler(CommandHandler("categories", categories))
    app.add_handler(CommandHandler("kurs", kurs_cmd))
    app.add_handler(CommandHandler("group_new", group_new))
    app.add_handler(CommandHandler("group_join", group_join))
    app.add_handler(CommandHandler("group_leave", group_leave))
    app.add_handler(CommandHandler("group_info", group_info))
    app.add_handler(CommandHandler("qarz_oldim", qarz_oldim_cmd))
    app.add_handler(CommandHandler("qarz_berdim", qarz_berdim_cmd))
    app.add_handler(CommandHandler("qarzlar", debts_list_cmd))
    app.add_handler(CommandHandler("qarz_yopish", qarz_yopish_cmd))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("excel", excel_export))
    app.add_handler(CallbackQueryHandler(debt_callback, pattern=r"^debtdir:|^debtlist$"))
    app.add_handler(CallbackQueryHandler(photo_amount_callback, pattern=r"^amtpick:|^amtmanual$"))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
