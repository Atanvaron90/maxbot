"""
MAX Bot — бот с поддержкой, информацией, условиями, рассылкой и проверкой долга.
Двухуровневая система обновления сверки из Google Таблиц:
  1) Триггеры (лист «Выгрузка продаж 5-ка») — ТОЛЬКО понедельник, раз в 15 мин,
     пока не обновятся в ОБОИХ таблицах. Зачёт (и рассылка всем) — только при
     ГЛУБОКОМ обновлении: у обеих таблиц изменилось число НЕПУСТЫХ строк.
  2) Лист «Сверка» — пн: раз в 15 мин, вт–пт: раз в 30 мин, сб–вс: не проверяется.
Контроль нового года: раз в неделю бот проверяет метку года в
«Исходные данные»!A2 первой таблицы; если метка устарела — предупреждение владельцу.
Администраторы: владелец (ADMIN_ID) + динамические (/addadmin, /deladmin, /adminlist,
кнопки ➕/➖ в меню владельца), хранятся в data/admins.json.
Команды: /start /menu, /help, /debt, /change_name, /cancel, /broadcast, /userinfo, /sentlog,
/refresh_debt, /debt_status, /addadmin, /deladmin, /adminlist.
"""

import os
import sys
import json
import hashlib
import asyncio
import logging
import re
import math
import time
import tempfile
import urllib.request
import urllib.error
from datetime import datetime, timedelta, timezone
from enum import Enum
from typing import Dict, Optional, List, Set

from dotenv import load_dotenv
from maxapi import Bot, Dispatcher, F
from maxapi.types import (
    MessageCreated, BotStarted, MessageCallback, Command,
    CallbackButton, ButtonsPayload, Attachment, BotCommand,
    InputMediaBuffer
)
from maxapi.enums.intent import Intent

try:
    import openpyxl
except ImportError:
    openpyxl = None
    logging.warning("openpyxl не установлен. Проверка долга не будет работать.")

try:
    from maxapi.enums import UploadType
except ImportError:
    try:
        from maxapi.enums.upload_type import UploadType
    except ImportError:
        UploadType = None

try:
    from zoneinfo import ZoneInfo
    MSK = ZoneInfo("Europe/Moscow")
except Exception:
    MSK = timezone(timedelta(hours=3), name="MSK")

load_dotenv()

DEBUG = os.getenv("DEBUG", "false").lower() == "true"

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

logging.basicConfig(
    level=logging.DEBUG if DEBUG else logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("MAX_BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))
ADMINS_FILE = os.getenv("ADMINS_FILE", "data/admins.json")
USERS_FILE = 'data/users.json'
# Журнал отправок бота (кому и какие сообщения отправлены): массовые рассылки
# и уведомления. Просмотр админом: команда /sentlog [N] или /sentlog clear.
SENT_LOG_FILE = os.getenv("SENT_LOG_FILE", "data/sent_log.json")
SENT_LOG_MAX = int(os.getenv("SENT_LOG_MAX", "1000"))

DEBT_SHEET_IDS: List[str] = [s.strip() for s in os.getenv("DEBT_SHEET_IDS", "").split(",") if s.strip()]
DEBT_DIR = os.getenv("DEBT_DIR", "data/debt")

DEBT_DATA_SHEET = os.getenv("DEBT_DATA_SHEET", "Сверка")
DEBT_SOURCE_SHEET = os.getenv("DEBT_SOURCE_SHEET", "Выгрузка продаж 5-ка")
# Столбец-триггер недельного обновления: заголовок в листе «Выгрузка продаж 5-ка».
# Уведомление о сверке уходит только при изменении данных в этом столбце
# (в обеих таблицах). Частичное заполнение других столбцов не считается.
TRIG_COLUMN_HEADER = os.getenv("TRIG_COLUMN_HEADER", "Количество")
_TRIG_COLUMN_ALIASES = {'количество', 'кол-во', 'колво'}
DEBT_YEAR_CELL = os.getenv("DEBT_YEAR_CELL", "Исходные данные!A2")
DEBT_YEAR_CHECK_WD = int(os.getenv("DEBT_YEAR_CHECK_WD", "0"))
DEBT_YEAR_CHECK_HOUR = int(os.getenv("DEBT_YEAR_CHECK_HOUR", "9"))

DEBT_CHECK_MONDAY_MINUTES = int(os.getenv("DEBT_CHECK_MONDAY_MINUTES", "15"))
DEBT_CHECK_WEEKDAY_MINUTES = int(os.getenv("DEBT_CHECK_WEEKDAY_MINUTES", "30"))
DEBT_MIN_DOWNLOAD_MINUTES = int(os.getenv("DEBT_MIN_DOWNLOAD_MINUTES", "10"))
DEBT_STATE_FILE = os.getenv("DEBT_STATE_FILE", "data/debt_state.json")

DOWNLOAD_TIMEOUT = int(os.getenv("DOWNLOAD_TIMEOUT", "120"))
DOWNLOAD_RETRIES = int(os.getenv("DOWNLOAD_RETRIES", "3"))

def _normalize_gdrive_url(url: str) -> str:
    """Преобразует любую ссылку Google Drive в прямую (без редиректа),
    которую сервер MAX скачивает сам без ошибки 'Failed to upload image'.
    Понимает форматы: /file/d/ID/view, /open?id=ID, /uc?...&id=ID.
    Уже прямые и сторонние (не Google) ссылки возвращает как есть."""
    if not url or "drive.google.com" not in url:
        return url
    m = re.search(r"drive\.google\.com/file/d/([A-Za-z0-9_-]{10,})", url)
    file_id = m.group(1) if m else None
    if not file_id:
        m = re.search(r"[?&]id=([A-Za-z0-9_-]{10,})", url)
        file_id = m.group(1) if m else None
    if not file_id:
        logger.warning(f"Не удалось извлечь FILE_ID из ссылки Google: {url}")
        return url
    direct = f"https://drive.usercontent.google.com/download?id={file_id}&export=download"
    if direct != url:
        logger.info(f"Ссылка Google преобразована в прямую: {url} → {direct}")
    return direct


# Картинки карточек: в .env можно вставлять ЛЮБУЮ ссылку Google Drive —
# она автоматически приводится к прямому виду при старте бота.
REGION1_APP = _normalize_gdrive_url(os.getenv("REGION1_APP", "https://example.com/app1.jpg"))
REGION1_QR = _normalize_gdrive_url(os.getenv("REGION1_QR", "https://example.com/qr1.jpg"))
REGION1_BANK = _normalize_gdrive_url(os.getenv("REGION1_BANK", "https://example.com/bank1.jpg"))
REGION2_APP = _normalize_gdrive_url(os.getenv("REGION2_APP", "https://example.com/app2.jpg"))
REGION2_QR = _normalize_gdrive_url(os.getenv("REGION2_QR", "https://example.com/qr2.jpg"))
REGION2_BANK = _normalize_gdrive_url(os.getenv("REGION2_BANK", "https://example.com/bank2.jpg"))

# Реквизиты для оплаты отправляются ТЕКСТОМ (не картинкой).
# У КАЖДОГО РЕГИОНА СВОЙ НЕЗАВИСИМЫЙ ОБРАЗЕЦ — правьте их по отдельности.
# Формат сообщения — markdown: жирный — **текст**, курсив — *текст*.
# Звёздочки-маркеры в чате НЕ отображаются, видно только начертание.
# ВАЖНО: если задаёте текст через .env (REGIONx_BANK_TEXT) — звёздочки
# нужно писать прямо в значении .env.
# Дополнительно можно переопределить текст через .env:
#   REGION1_BANK_TEXT — Краснодар, REGION2_BANK_TEXT — КМВ, КЧР.

# --- Регион 1: Краснодар ---
BANK_DETAILS_1 = (
    "**Индивидуальный предприниматель Ершов Максим Сергеевич**\n"
    "**ИНН:** 370225603829\n"
    "**ОГРНИП:** 322370200022099\n"
    "**Номер счета:** 40802810138120000533\n"
    "**Банк:** Ивановский РФ АО \"Россельхозбанк\"\n"
    "**БИК:** 044525780\n"
    "**Корр. счет:** 30101810745374525780"
)

# --- Регион 2: КМВ, КЧР ---
BANK_DETAILS_2 = (
    "**Индивидуальный предприниматель Ершов Максим Сергеевич**\n"
    "**ИНН:** 370225603829\n"
    "**ОГРНИП:** 322370200022099\n"
    "**Номер счета:** 40802810138120000533\n"
    "**Банк:** Ивановский РФ АО \"Россельхозбанк\"\n"
    "**БИК:** 044525780\n"
    "**Корр. счет:** 30101810745374525780"
)

REGION1_BANK_TEXT = os.getenv("REGION1_BANK_TEXT", BANK_DETAILS_1)
REGION2_BANK_TEXT = os.getenv("REGION2_BANK_TEXT", BANK_DETAILS_2)

if not BOT_TOKEN:
    logger.error("MAX_BOT_TOKEN не найден в .env")
    exit(1)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

KNOWN_COMMANDS = {'/start', '/menu', '/help', '/debt', '/change_name', '/cancel',
                  '/broadcast', '/userinfo', '/sentlog',
                  '/refresh_debt', '/debt_status',
                  '/addadmin', '/deladmin', '/adminlist'}


class UserState(Enum):
    IDLE = "idle"
    BROADCAST_MODE = "broadcast"
    DEBT_NAME_MODE = "debt_name"
    DEBT_CLARIFY_MODE = "debt_clarify"
    WAITING_FOR_FORWARD = "waiting_for_forward"
    ADMIN_DEBT_MODE = "admin_debt"
    ADMIN_DEBT_CLARIFY_MODE = "admin_debt_clarify"
    WAITING_ADDADMIN_ID = "waiting_addadmin_id"
    WAITING_DELADMIN_ID = "waiting_deladmin_id"


user_states: Dict[int, UserState] = {}
user_clarify_data: Dict[int, Dict[str, object]] = {}
admin_debt_data: Dict[int, Dict[str, object]] = {}

_debt_cache: Dict = {'mtime': None, 'rows': None, 'loaded_at': None}

_debt_state: Dict = {
    'trig_baseline': None,
    'data_baseline': None,
    'trig_rows_baseline': None,
    'trig_committed_on': None,
    'year_warned_for': None,
    'year_label': None,
    'last_sigs_t': None,
    'last_sigs_d': None,
    'last_rows_t': None,
    'last_check': None,
}


# ---------- АДМИНИСТРАТОРЫ ----------

_dynamic_admins: Set[int] = set()


def _load_admins():
    global _dynamic_admins
    try:
        os.makedirs('data', exist_ok=True)
        if os.path.exists(ADMINS_FILE):
            with open(ADMINS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            _dynamic_admins = {int(x) for x in data.get('admins', []) if str(x).lstrip('-').isdigit()}
    except Exception as e:
        logger.error(f"Ошибка загрузки админов: {e}")
        _dynamic_admins = set()


def _save_admins():
    try:
        os.makedirs('data', exist_ok=True)
        with open(ADMINS_FILE, 'w', encoding='utf-8') as f:
            json.dump({'admins': sorted(_dynamic_admins)}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Ошибка сохранения админов: {e}")


_load_admins()


def is_admin(user_id: int) -> bool:
    return user_id == ADMIN_ID or user_id in _dynamic_admins


def is_owner(user_id: int) -> bool:
    return user_id == ADMIN_ID


# ---------- РАБОТА С БАЗОЙ ПОЛЬЗОВАТЕЛЕЙ ----------

def load_users() -> Dict:
    try:
        os.makedirs('data', exist_ok=True)
        if os.path.exists(USERS_FILE):
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        logger.error(f"Ошибка загрузки: {e}")
    return {}


def save_users(users: Dict):
    try:
        os.makedirs('data', exist_ok=True)
        with open(USERS_FILE, 'w', encoding='utf-8') as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Ошибка сохранения: {e}")


def add_user(user_id: int, chat_id: int, username: str = None, name: str = None, full_name: str = None):
    users = load_users()
    uid = str(user_id)
    is_new = uid not in users
    now_msk = datetime.now(MSK).isoformat()

    if is_new:
        users[uid] = {
            'chat_id': chat_id,
            'username': username,
            'name': name,
            'full_name': full_name,
            'created_at': now_msk,
            'last_activity': now_msk
        }
    else:
        existing = users[uid]
        if chat_id is not None:
            existing['chat_id'] = chat_id
        if username is not None:
            existing['username'] = username
        if name is not None:
            existing['name'] = name
        if full_name is not None:
            existing['full_name'] = full_name
        existing['last_activity'] = now_msk
        if 'created_at' not in existing:
            existing['created_at'] = now_msk
        users[uid] = existing

    save_users(users)
    logger.info(f"Обновлён пользователь: id={user_id}, is_new={is_new}")
    return is_new


def ensure_user(user_id: int, chat_id: int, username: str = None, name: str = None):
    return add_user(user_id, chat_id, username, name)


def set_user_full_name(user_id: int, full_name: str):
    users = load_users()
    uid = str(user_id)
    if uid in users:
        users[uid]['full_name'] = full_name
        save_users(users)
        logger.info(f"Сохранено ФИО для user_id={user_id}")
        return True
    logger.warning(f"Не удалось сохранить ФИО для user_id={user_id}: пользователь не найден")
    return False


def clear_user_full_name(user_id: int):
    users = load_users()
    uid = str(user_id)
    if uid in users:
        users[uid]['full_name'] = None
        save_users(users)
        logger.info(f"Очищено ФИО для user_id={user_id}")
        return True
    return False


# ---------- МЕНЮ ----------

def create_main_menu(user_id: int) -> Attachment:
    admin = is_admin(user_id)
    row1 = [
        CallbackButton(text="📋 Условия", payload="conditions", intent=Intent.DEFAULT),
        CallbackButton(text="💳 Долг", payload="debt", intent=Intent.DEFAULT)
    ]
    row2 = [
        CallbackButton(text="ℹ️ О боте", payload="info", intent=Intent.DEFAULT),
        CallbackButton(text="❓ Справка", payload="help_command", intent=Intent.DEFAULT)
    ]
    rows = [row1, row2]

    if admin:
        row3: List[CallbackButton] = []
        if admin:
            row3.append(CallbackButton(text="📢 Новости", payload="broadcast", intent=Intent.DEFAULT))
        row3.append(CallbackButton(text="👤 ID контакта", payload="userinfo", intent=Intent.DEFAULT))
        rows.append(row3)
        rows.append([
            CallbackButton(text="🔄 Обновить сверку", payload="refresh_debt", intent=Intent.DEFAULT),
            CallbackButton(text="📊 Статус сверки", payload="debt_status", intent=Intent.DEFAULT)
        ])
        if is_owner(user_id):
            rows.append([
                CallbackButton(text="👥 Админы", payload="adminlist", intent=Intent.DEFAULT)
            ])
            rows.append([
                CallbackButton(text="➕ Админ", payload="addadmin", intent=Intent.DEFAULT),
                CallbackButton(text="➖ Админ", payload="deladmin", intent=Intent.DEFAULT)
            ])

    buttons = ButtonsPayload(buttons=rows)
    return Attachment(type="inline_keyboard", payload=buttons)


def create_debt_result_menu() -> Attachment:
    buttons = ButtonsPayload(buttons=[
        [CallbackButton(text="🏠 Главное меню", payload="back_to_main", intent=Intent.DEFAULT)],
        [CallbackButton(text="🔄 Сменить ФИО", payload="change_name", intent=Intent.DEFAULT)]
    ])
    return Attachment(type="inline_keyboard", payload=buttons)


def create_admin_debt_menu() -> Attachment:
    buttons = ButtonsPayload(buttons=[
        [CallbackButton(text="🔍 Проверить другое ФИО", payload="admin_debt_again", intent=Intent.DEFAULT)],
        [CallbackButton(text="🏠 Главное меню", payload="back_to_main", intent=Intent.DEFAULT)],
    ])
    return Attachment(type="inline_keyboard", payload=buttons)


def create_regions_menu() -> Attachment:
    buttons = ButtonsPayload(buttons=[
        [
            CallbackButton(text="📍 Краснодар", payload="region_1", intent=Intent.DEFAULT),
            CallbackButton(text="📍 КМВ, КЧР", payload="region_2", intent=Intent.DEFAULT)
        ],
        [
            CallbackButton(text="🔙 Назад", payload="back_to_main", intent=Intent.DEFAULT)
        ]
    ])
    return Attachment(type="inline_keyboard", payload=buttons)


def create_materials_menu(region: str) -> Attachment:
    buttons = ButtonsPayload(buttons=[
        [CallbackButton(text="📄 Заявление", payload=f"{region}_app", intent=Intent.DEFAULT)],
        [CallbackButton(text="📱 QR-код", payload=f"{region}_qr", intent=Intent.DEFAULT)],
        [CallbackButton(text="💳 Реквизиты", payload=f"{region}_bank", intent=Intent.DEFAULT)],
        [CallbackButton(text="🔙 Назад к регионам", payload="back_to_regions", intent=Intent.DEFAULT)]
    ])
    return Attachment(type="inline_keyboard", payload=buttons)


def create_cancel_menu() -> Attachment:
    buttons = ButtonsPayload(buttons=[
        [CallbackButton(text="❌ Отмена", payload="cancel_action", intent=Intent.DEFAULT)]
    ])
    return Attachment(type="inline_keyboard", payload=buttons)


# ---------- ФУНКЦИИ ДЛЯ РАБОТЫ С ДОЛГОМ ----------

def parse_balance(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    s = re.sub(r'[^\d,.\-]', '', str(value))
    if not s.strip('-'):
        return 0.0

    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    else:
        s = s.replace(',', '.')
        if s.count('.') > 1:
            head, _, tail = s.rpartition('.')
            s = head.replace('.', '') + '.' + tail

    try:
        return float(s)
    except ValueError:
        logger.warning(f"Не удалось преобразовать значение в число: {value!r}")
        return 0.0


def _norm_word(w: str) -> str:
    return w.lower().replace('ё', 'е')


def clean_name(name: str) -> str:
    name = name.strip()
    if '(' in name:
        name = name.split('(')[0].strip()
    if ',' in name:
        name = name.split(',')[0].strip()
    return name


def _norm_sheet_name(name: str) -> str:
    s = str(name).replace('\xa0', ' ').replace('\u2007', ' ').replace('\u202f', ' ')
    s = re.sub(r'\s+', ' ', s).strip().lower()
    for dash in ('\u2010', '\u2011', '\u2012', '\u2013', '\u2014', '\u2212'):
        s = s.replace(dash, '-')
    return s


def _debt_paths() -> List[str]:
    return [os.path.join(DEBT_DIR, f"sheet_{i}.xlsx") for i in range(1, len(DEBT_SHEET_IDS) + 1)]


def _cache_fresh() -> bool:
    paths = _debt_paths()
    if not paths or any(not os.path.exists(p) for p in paths):
        return False
    try:
        oldest = min(os.path.getmtime(p) for p in paths)
    except OSError:
        return False
    return (time.time() - oldest) < DEBT_MIN_DOWNLOAD_MINUTES * 60


def _load_debt_rows() -> Optional[List[Dict]]:
    if not openpyxl:
        logger.error("openpyxl не установлен")
        return None

    paths = _debt_paths()
    if not paths:
        logger.error("Список таблиц пуст (DEBT_SHEET_IDS)")
        return None

    existing = [p for p in paths if os.path.exists(p)]
    if not existing:
        logger.error("Ни одного кэш-файла сверки не найдено")
        return None

    try:
        mtimes = tuple(os.path.getmtime(p) for p in existing)
    except OSError as e:
        logger.error(f"Не удалось получить mtime: {e}")
        return None

    if _debt_cache['mtime'] == mtimes and _debt_cache['rows'] is not None:
        return _debt_cache['rows']

    last_error = None
    for attempt in range(3):
        try:
            rows: List[Dict] = []
            target_norm = _norm_sheet_name(DEBT_DATA_SHEET)
            for path in existing:
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                sheet_name = next((s for s in wb.sheetnames if _norm_sheet_name(s) == target_norm), None)
                if sheet_name is None:
                    logger.warning(
                        f"В файле {os.path.basename(path)} нет листа «{DEBT_DATA_SHEET}». "
                        f"Реальные имена: {[repr(s) for s in wb.sheetnames]}"
                    )
                    wb.close()
                    continue
                sheet = wb[sheet_name]
                for row in sheet.iter_rows(min_row=2, max_col=5):
                    cell_name = row[0].value
                    if not cell_name:
                        continue
                    rows.append({
                        'name': str(cell_name).strip(),
                        'balance': parse_balance(row[4].value),
                        'sheet': sheet_name,
                        'file': os.path.basename(path),
                    })
                wb.close()

            _debt_cache['mtime'] = mtimes
            _debt_cache['rows'] = rows
            _debt_cache['loaded_at'] = datetime.now(MSK)
            logger.info(f"Лист «{DEBT_DATA_SHEET}» прочитан: {len(rows)} строк из {len(existing)} файл(ов)")
            return rows
        except Exception as e:
            last_error = e
            logger.warning(f"Попытка {attempt + 1}/3 прочитать файлы долгов не удалась: {e}")
            time.sleep(0.5)

    logger.error(f"Не удалось прочитать файлы долгов после 3 попыток: {last_error}")
    if _debt_cache['rows'] is not None:
        logger.warning("Использую предыдущую версию данных сверки")
        return _debt_cache['rows']
    return None


def search_debt_by_name(full_name: str) -> Optional[List[Dict]]:
    rows = _load_debt_rows()
    if rows is None:
        return None

    words = full_name.strip().split()
    if len(words) < 2:
        return []

    n = min(len(words), 3)
    search_words = [_norm_word(w) for w in words[:n]]
    matches = []

    for item in rows:
        cell_words = clean_name(item['name']).split()
        if len(cell_words) < n:
            continue
        cell_search = [_norm_word(w) for w in cell_words[:n]]
        if search_words == cell_search:
            matches.append(item)

    return matches


def group_matches(matches: List[Dict]) -> Dict[str, List[Dict]]:
    groups: Dict[str, List[Dict]] = {}
    for item in matches:
        clean = clean_name(item['name'])
        key = ' '.join(_norm_word(w) for w in clean.split()[:3])
        groups.setdefault(key, []).append(item)
    return groups


def get_last_sunday() -> str:
    now = datetime.now(MSK)
    today = now.date()
    if today.weekday() == 0 and now.hour < 13:
        delta = timedelta(days=8)
    else:
        if today.weekday() == 6:
            delta = timedelta(days=7)
        else:
            days_to_subtract = today.weekday() + 1
            delta = timedelta(days=days_to_subtract)
    last_sunday = today - delta
    return last_sunday.strftime("%d.%m.%Y")


def format_debt_result(matches: List[Dict], date_str: str) -> str:
    groups = group_matches(matches)

    if len(groups) > 1:
        all_negative = [item for items in groups.values() for item in items if item['balance'] < 0]
        if all_negative:
            lines = [f"⚠️ **Найдены задолженности по разным магазинам (по {date_str} включительно):**"]
            for idx, item in enumerate(all_negative, 1):
                rounded = math.ceil(abs(item['balance']))
                lines.append(f"{idx}. {item['name']} — задолженность: {rounded} руб.")
            lines.append("\n_При необходимости можете уточнить информацию у менеджера._")
            return "\n".join(lines)
        return f"✅ **По {date_str} включительно у вас нет задолженности.**"

    group_items = next(iter(groups.values()))
    negative_items = [item for item in group_items if item['balance'] < 0]

    if not negative_items:
        return f"✅ **По {date_str} включительно у вас нет задолженности.**"
    elif len(negative_items) == 1:
        rounded = math.ceil(abs(negative_items[0]['balance']))
        return f"⚠️ **По {date_str} включительно у вас есть задолженность на сумму {rounded} руб.**"
    else:
        lines = [f"⚠️ **Найдены задолженности по разным магазинам (по {date_str} включительно):**"]
        for idx, item in enumerate(negative_items, 1):
            rounded = math.ceil(abs(item['balance']))
            lines.append(f"{idx}. {item['name']} — задолженность: {rounded} руб.")
        lines.append("\n_При необходимости можете уточнить информацию у менеджера._")
        return "\n".join(lines)


async def send_debt_result(chat_id: int, user_id: int, matches: List[Dict],
                           menu: Optional[Attachment] = None):
    date_str = get_last_sunday()
    result_text = format_debt_result(matches, date_str)
    await bot.send_message(
        chat_id=chat_id,
        text=result_text,
        format="markdown",
        attachments=[menu if menu is not None else create_debt_result_menu()]
    )


# ---------- СИСТЕМА ОБНОВЛЕНИЯ СВЕРКИ ----------

def _download_debt_file_sync(url: str, tmp_path: str) -> None:
    last_error: Exception = None
    for attempt in range(1, DOWNLOAD_RETRIES + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
                data = resp.read()

            if not data.startswith(b"PK"):
                raise ValueError("Ответ не является xlsx-файлом (проверьте доступ к таблице)")

            with open(tmp_path, "wb") as f:
                f.write(data)
            return

        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise ValueError(
                    f"HTTP {e.code}: таблица закрыта доступом. Откройте «Настройки доступа» → "
                    "«Все, у кого есть ссылка» → Читатель."
                ) from e
            if e.code == 404:
                raise ValueError("HTTP 404: таблица не найдена — проверьте DEBT_SHEET_IDS.") from e
            last_error = e
            logger.warning(f"Скачивание, попытка {attempt}/{DOWNLOAD_RETRIES}: HTTP {e.code}")
        except Exception as e:
            last_error = e
            logger.warning(f"Скачивание, попытка {attempt}/{DOWNLOAD_RETRIES} не удалась: {e}")

        if attempt < DOWNLOAD_RETRIES:
            time.sleep(5 * attempt)

    raise last_error


def _cleanup_files(paths: List[str]):
    for p in paths:
        try:
            if os.path.exists(p):
                os.remove(p)
        except OSError:
            pass


async def _download_tables_tmp() -> Optional[List[str]]:
    tmps = [p.replace(".xlsx", ".tmp.xlsx") for p in _debt_paths()]
    os.makedirs(DEBT_DIR, exist_ok=True)

    try:
        for sid, path in zip(DEBT_SHEET_IDS, tmps):
            url = (f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
                   f"&nocache={int(time.time() * 1000)}")
            await asyncio.to_thread(_download_debt_file_sync, url, path)
            logger.info(f"Таблица {sid[:10]}… скачана")
        return tmps
    except Exception as e:
        logger.error(f"Не удалось скачать таблицу: {e}")
        _cleanup_files(tmps)
        return None


def _sheet_signature(path: str, sheet_name: str) -> Optional[str]:
    """md5-отпечаток содержимого листа. Хешируется лист ПОЛНОСТЬЮ."""
    if not openpyxl:
        return None

    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        logger.warning(f"openpyxl (read_only) не смог открыть {os.path.basename(path)}: {e}")
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e2:
            logger.warning(f"Не удалось открыть {os.path.basename(path)} совсем: {e2}")
            return None

    try:
        target = _norm_sheet_name(sheet_name)
        name = next((s for s in wb.sheetnames if _norm_sheet_name(s) == target), None)
        if name is None:
            logger.warning(
                f"Лист «{sheet_name}» не найден в {os.path.basename(path)}. "
                f"Реальные имена листов: {[repr(s) for s in wb.sheetnames]}"
            )
            return None

        h = hashlib.md5()
        for row in wb[name].iter_rows(values_only=True):
            h.update(repr(row).encode('utf-8'))
        return h.hexdigest()
    finally:
        wb.close()


def _column_signature(path: str, sheet_name: str) -> Optional[str]:
    """md5-отпечаток столбца «Количество» листа-триггера.
    Заголовок ищется в первых 5 строках листа; хешируются только НЕПУСТЫЕ
    значения столбца ниже заголовка, плюс их количество. Пустые строки-хвост
    и изменения в других столбцах отпечаток НЕ меняют.
    None — лист или столбец не найден."""
    if not openpyxl:
        return None

    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        logger.warning(f"openpyxl (read_only) не смог открыть {os.path.basename(path)}: {e}")
        return None

    try:
        target = _norm_sheet_name(sheet_name)
        nm = next((s for s in wb.sheetnames if _norm_sheet_name(s) == target), None)
        if nm is None:
            logger.warning(
                f"Лист «{sheet_name}» не найден в {os.path.basename(path)}. "
                f"Реальные имена листов: {[repr(s) for s in wb.sheetnames]}"
            )
            return None

        header_norm = _norm_word(TRIG_COLUMN_HEADER)
        target_col = None
        scan_rows = 0
        values: List[str] = []

        for row in wb[nm].iter_rows(values_only=True):
            scan_rows += 1

            # Пока не найден заголовок — ищем его (первые 5 строк листа)
            if target_col is None:
                for ci, v in enumerate(row, start=1):
                    if v is None:
                        continue
                    nv = _norm_word(str(v)).strip()
                    if nv in _TRIG_COLUMN_ALIASES:
                        target_col = ci
                        break
                if target_col is None and scan_rows >= 5:
                    logger.warning(
                        f"Столбец «{TRIG_COLUMN_HEADER}» не найден в первых 5 строках "
                        f"листа «{sheet_name}» ({os.path.basename(path)})")
                    return None
                continue

            # Заголовок найден — собираем значения столбца
            v = row[target_col - 1] if len(row) >= target_col else None
            if v is not None and str(v).strip() != '':
                values.append(repr(v))

        if target_col is None:
            logger.warning(f"Столбец «{TRIG_COLUMN_HEADER}» не найден в листе «{sheet_name}»")
            return None

        h = hashlib.md5()
        for v in values:
            h.update(v.encode('utf-8'))
        h.update(f"n={len(values)}".encode('utf-8'))
        return h.hexdigest()
    finally:
        wb.close()


def _trig_signatures(tmps: List[str]) -> List[Optional[str]]:
    """Сигнатуры столбца «Количество» листа-триггера для каждой таблицы."""
    return [_column_signature(t, DEBT_SOURCE_SHEET) for t in tmps]
    

def _signatures(tmps: List[str], sheet_name: str) -> List[Optional[str]]:
    return [_sheet_signature(t, sheet_name) for t in tmps]


def _rows_counts(paths: List[str], sheet_name: str) -> List[Optional[int]]:
    """Число НЕПУСТЫХ строк листа в каждом файле (глубокая проверка триггеров)."""
    out: List[Optional[int]] = []
    for p in paths:
        try:
            wb2 = openpyxl.load_workbook(p, data_only=True, read_only=True)
        except Exception:
            out.append(None)
            continue
        try:
            target2 = _norm_sheet_name(sheet_name)
            nm = next((s for s in wb2.sheetnames if _norm_sheet_name(s) == target2), None)
            if nm is None:
                out.append(None)
                continue
            cnt = 0
            for row in wb2[nm].iter_rows(values_only=True):
                if any(v is not None and str(v).strip() != '' for v in row):
                    cnt += 1
            out.append(cnt)
        finally:
            wb2.close()
    return out


def _parse_cell_addr(addr: str) -> Optional[tuple]:
    """Разбирает адрес ячейки 'A2' -> (строка=2, столбец=1). None — не удалось."""
    m = re.match(r'^([A-Za-z]+)(\d+)$', addr.strip())
    if not m:
        return None
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return (int(m.group(2)), col)


def _read_year_cell(path: str, cell_spec: str) -> Optional[str]:
    """YEAR: читает ячейку-метку года. cell_spec = 'Лист!A2' — читается
    именно указанная ячейка. Числовые значения приводятся к целому виду
    (2026.0 -> '2026'). None — не удалось прочитать."""
    try:
        sheet_name, addr = cell_spec.split("!", 1)
    except ValueError:
        return None

    rc = _parse_cell_addr(addr)
    if rc is None:
        logger.warning(f"Не удалось разобрать адрес ячейки: {cell_spec!r}")
        return None
    target_row, target_col = rc

    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None
    try:
        nm = next((s for s in wb.sheetnames if _norm_sheet_name(s) == _norm_sheet_name(sheet_name)), None)
        if nm is None:
            return None
        val = None
        r_idx = 0
        for row in wb[nm].iter_rows(min_row=1, max_row=target_row,
                                    min_col=target_col, max_col=target_col,
                                    values_only=True):
            r_idx += 1
            if r_idx == target_row:
                val = row[0] if row else None
        if val is None:
            return ""
        # Число из Excel часто приходит как float (2026.0) — приводим к '2026'
        if isinstance(val, float) and val.is_integer():
            return str(int(val))
        return str(val).strip()
    finally:
        wb.close()


def _commit_tables(tmps: List[str]):
    for tmp, path in zip(tmps, _debt_paths()):
        os.replace(tmp, path)
    logger.info("Кэш сверки обновлён")


def _load_debt_state():
    try:
        if os.path.exists(DEBT_STATE_FILE):
            with open(DEBT_STATE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            n = len(DEBT_SHEET_IDS)
            tb = data.get('trig_baseline')
            db = data.get('data_baseline')
            if isinstance(tb, list) and len(tb) == n:
                _debt_state['trig_baseline'] = tb
            if isinstance(db, list) and len(db) == n:
                _debt_state['data_baseline'] = db
            trb = data.get('trig_rows_baseline')
            if isinstance(trb, list) and len(trb) == n:
                _debt_state['trig_rows_baseline'] = trb
            tco = data.get('trig_committed_on')
            if isinstance(tco, str):
                _debt_state['trig_committed_on'] = tco
            yw = data.get('year_warned_for')
            if isinstance(yw, str):
                _debt_state['year_warned_for'] = yw
            yl = data.get('year_label')
            if isinstance(yl, str):
                _debt_state['year_label'] = yl
    except Exception as e:
        logger.warning(f"Не удалось прочитать состояние сверки: {e}")


def _save_debt_state():
    try:
        os.makedirs('data', exist_ok=True)
        with open(DEBT_STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({
                'trig_baseline': _debt_state['trig_baseline'],
                'data_baseline': _debt_state['data_baseline'],
                'trig_rows_baseline': _debt_state.get('trig_rows_baseline'),
                'trig_committed_on': _debt_state['trig_committed_on'],
                'year_warned_for': _debt_state.get('year_warned_for'),
                'year_label': _debt_state.get('year_label'),
            }, f, ensure_ascii=False)
    except Exception as e:
        logger.error(f"Не удалось сохранить состояние сверки: {e}")


# ---------- ЖУРНАЛ ОТПРАВОК БОТА ----------

def _load_sent_log() -> list:
    """Читает журнал отправок (SENT_LOG_FILE). Пустой список при ошибке/отсутствии."""
    try:
        if os.path.exists(SENT_LOG_FILE):
            with open(SENT_LOG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception as e:
        logger.error(f"Не удалось прочитать журнал отправок: {e}")
    return []


def _log_sent_bulk(recipients: List[tuple], source: str, text: str):
    """Добавляет в журнал факты УСПЕШНЫХ массовых отправок ботом.
    recipients — список пар (user_id, user_data); текст обрезается до 200 симв.
    Хранятся последние SENT_LOG_MAX записей (ротация старых)."""
    if not recipients:
        return
    try:
        os.makedirs('data', exist_ok=True)
        entries = _load_sent_log()
        ts = datetime.now(MSK).strftime('%d.%m.%Y %H:%M:%S')
        for uid, udata in recipients:
            name = '—'
            chat_id_u = None
            if isinstance(udata, dict):
                name = udata.get('full_name') or udata.get('name') or '—'
                chat_id_u = udata.get('chat_id')
            entries.append({
                'ts': ts,
                'user_id': uid,
                'chat_id': chat_id_u,
                'name': name,
                'source': source,
                'text': (text or '')[:200],
            })
        if len(entries) > SENT_LOG_MAX:
            entries = entries[-SENT_LOG_MAX:]
        with open(SENT_LOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(entries, f, ensure_ascii=False, indent=1)
    except Exception as e:
        logger.error(f"Не удалось записать журнал отправок: {e}")


async def notify_debt_updated():
    text = "Сверка обновлена. Можно проверить наличие актуальной задолженности."
    users = load_users()
    sent = 0
    errors = 0
    ok_recipients: List[tuple] = []
    for uid, user_data in users.items():
        chat_id_user = user_data.get('chat_id')
        if not chat_id_user:
            continue
        try:
            await bot.send_message(chat_id=chat_id_user, text=text)
            sent += 1
            ok_recipients.append((uid, user_data))
        except Exception as e:
            errors += 1
            logger.error(f"Уведомление не доставлено пользователю {uid}: {e}")
        await asyncio.sleep(0.2)
    logger.info(f"Уведомление об обновлении сверки: отправлено {sent}, ошибок {errors}")
    _log_sent_bulk(ok_recipients, "Уведомление о сверке", text)


def _rows_gain_text(st: Dict, cur_rows: List[Optional[int]]) -> Optional[str]:
    base_rows = st.get('trig_rows_baseline')
    if not base_rows or len(base_rows) != len(cur_rows):
        return None
    gains = []
    for i in range(len(cur_rows)):
        if cur_rows[i] is not None and base_rows[i] is not None and cur_rows[i] != base_rows[i]:
            gains.append(f"табл.{i+1}: {cur_rows[i]} (было {base_rows[i]})")
    if not gains:
        return None
    return "📈 Выгрузки: " + "; ".join(gains)


# ---------- АВТОВОССТАНОВЛЕНИЕ КЭША ПРИ ЗАПРОСЕ ДОЛГА ----------

_debt_recover_lock: Optional[asyncio.Lock] = None


async def ensure_debt_cache() -> bool:
    global _debt_recover_lock
    if _debt_recover_lock is None:
        _debt_recover_lock = asyncio.Lock()

    async with _debt_recover_lock:
        if _load_debt_rows() is not None:
            return True

        logger.warning("Кэш сверки недоступен — скачиваю таблицы по запросу")
        tmps = await _download_tables_tmp()
        if tmps is None:
            logger.error("Восстановить кэш сверки не удалось (таблицы не скачались)")
            return False

        trig_sigs = await asyncio.to_thread(_trig_signatures, tmps)
        data_sigs = await asyncio.to_thread(_signatures, tmps, DEBT_DATA_SHEET)

        _commit_tables(tmps)

        if all(s is not None for s in data_sigs):
            _debt_state['data_baseline'] = data_sigs[:]
            if _debt_state['trig_baseline'] is None and all(s is not None for s in trig_sigs):
                _debt_state['trig_baseline'] = trig_sigs[:]
                _debt_state['trig_rows_baseline'] = await asyncio.to_thread(
                    _rows_counts, tmps, DEBT_SOURCE_SHEET)
        _save_debt_state()

        _load_debt_rows()
        logger.info("Кэш сверки восстановлен по запросу")
        return True


async def search_debt_with_recovery(full_name: str) -> Optional[List[Dict]]:
    matches = await asyncio.to_thread(search_debt_by_name, full_name)
    if matches is None:
        if await ensure_debt_cache():
            matches = await asyncio.to_thread(search_debt_by_name, full_name)
    return matches


# ---------- ПЛАНОВЫЕ ПРОВЕРКИ ----------

async def startup_check() -> str:
    """Проверка при старте бота. Триггер = изменение столбца «Количество»
    в ОБОИХ таблицах, зачёт с рассылкой — только в понедельник; вне понедельника
    расхождение baseline — тихая синхронизация. Меняется только «Сверка» — тихо."""
    _load_debt_state()
    st = _debt_state
    if _cache_fresh() and st['trig_baseline'] and st['data_baseline']:
        return (f"✅ Кэш сверки свежий (< {DEBT_MIN_DOWNLOAD_MINUTES} мин) — "
                "стартовое скачивание пропущено.")

    n = len(DEBT_SHEET_IDS)

    tmps = await _download_tables_tmp()
    if tmps is None:
        return "❌ Не удалось скачать одну из таблиц (см. лог)."

    trig_sigs = await asyncio.to_thread(_trig_signatures, tmps)
    if any(s is None for s in trig_sigs):
        _cleanup_files(tmps)
        return (f"⚠️ Не найден столбец «{TRIG_COLUMN_HEADER}» в листе "
                f"«{DEBT_SOURCE_SHEET}» (подробности в логе).")

    data_sigs = await asyncio.to_thread(_signatures, tmps, DEBT_DATA_SHEET)
    if any(s is None for s in data_sigs):
        _cleanup_files(tmps)
        return (f"⚠️ Не найден лист «{DEBT_DATA_SHEET}». "
                "Точные имена листов см. в предупреждениях выше в логе.")

    cur_rows = await asyncio.to_thread(_rows_counts, tmps, DEBT_SOURCE_SHEET)
    year_label = await asyncio.to_thread(_read_year_cell, tmps[0], DEBT_YEAR_CELL)

    st = _debt_state
    st['last_sigs_t'] = trig_sigs
    st['last_sigs_d'] = data_sigs
    st['last_rows_t'] = cur_rows[:]
    st['year_label'] = year_label
    st['last_check'] = datetime.now(MSK)

    if st['trig_baseline'] is None or st['data_baseline'] is None:
        _commit_tables(tmps)
        st['trig_baseline'] = trig_sigs[:]
        st['data_baseline'] = data_sigs[:]
        st['trig_rows_baseline'] = cur_rows[:]
        _save_debt_state()
        rows = _load_debt_rows()
        count = len(rows) if rows is not None else "?"
        return f"✅ Сверка загружена ({count} записей). Базовое состояние зафиксировано."

    now = datetime.now(MSK)
    today = now.date().isoformat()
    today_wd = now.weekday()

    trig_fingerprint_changed = any(trig_sigs[i] != st['trig_baseline'][i] for i in range(n))

    # Зачёт обновления за время простоя — ТОЛЬКО в понедельник
    if trig_fingerprint_changed and today_wd == 0:
        _commit_tables(tmps)
        st['trig_baseline'] = trig_sigs[:]
        st['data_baseline'] = data_sigs[:]
        st['trig_rows_baseline'] = cur_rows[:]
        st['trig_committed_on'] = today
        _save_debt_state()
        rows = _load_debt_rows()
        count = len(rows) if rows is not None else "?"
        await notify_debt_updated()
        gain = _rows_gain_text(st, cur_rows) or ""
        return (f"✅ За время отключения столбец «{TRIG_COLUMN_HEADER}» обновился в обеих "
                f"таблицах. Сверка актуальна ({count} записей), уведомление отправлено. {gain}")

    # Вне понедельника расхождение триггеров — тихая синхронизация
    if trig_fingerprint_changed:
        st['trig_baseline'] = trig_sigs[:]
        st['trig_rows_baseline'] = cur_rows[:]
        _save_debt_state()
        logger.info("Триггеры: baseline синхронизирован вне понедельника (без рассылки)")

    if any(data_sigs[i] != st['data_baseline'][i] for i in range(n)):
        _commit_tables(tmps)
        st['data_baseline'] = data_sigs[:]
        _save_debt_state()
        rows = _load_debt_rows()
        count = len(rows) if rows is not None else "?"
        gain = _rows_gain_text(st, cur_rows)
        extra = f" {gain}" if gain else ""
        return (f"🔄 Лист «{DEBT_DATA_SHEET}» обновился — сверка актуализирована "
                f"({count} записей), без рассылки.{extra}")

    _cleanup_files(tmps)
    rows = _load_debt_rows()
    count = len(rows) if rows is not None else "?"
    return f"✅ Сверка актуальна ({count} записей), изменений нет."


async def scheduled_check(force_weekday: bool = False) -> str:
    """Плановая проверка.
    1) Триггеры: только в понедельник (или force_weekday для ручного /refresh_debt)
       и только если ещё не зачтены сегодня. Триггер = изменение данных в столбце
       «Количество» листа «Выгрузка продаж 5-ка» в ОБОИХ таблицах → зачёт + рассылка.
    2) Сверка: пн–пт; изменилась → тихая подмена кэша."""
    n = len(DEBT_SHEET_IDS)
    if not force_weekday and _cache_fresh():
        return "— Кэш сверки свежий, скачивание пропущено."

    tmps = await _download_tables_tmp()
    if tmps is None:
        return "❌ Не удалось скачать одну из таблиц (см. лог). Пробую при следующей проверке."

    # Триггер: отпечатки столбца «Количество»
    trig_sigs = await asyncio.to_thread(_trig_signatures, tmps)
    if any(s is None for s in trig_sigs):
        _cleanup_files(tmps)
        return (f"⚠️ Не найден столбец «{TRIG_COLUMN_HEADER}» в листе "
                f"«{DEBT_SOURCE_SHEET}» (подробности в логе). Зачёт не выполнялся.")

    data_sigs = await asyncio.to_thread(_signatures, tmps, DEBT_DATA_SHEET)
    if any(s is None for s in data_sigs):
        _cleanup_files(tmps)
        return (f"⚠️ Не найден лист «{DEBT_DATA_SHEET}». "
                "Точные имена листов см. в предупреждениях выше в логе.")

    cur_rows = await asyncio.to_thread(_rows_counts, tmps, DEBT_SOURCE_SHEET)
    year_label = await asyncio.to_thread(_read_year_cell, tmps[0], DEBT_YEAR_CELL)

    now = datetime.now(MSK)
    today = now.date().isoformat()
    wd = now.weekday()

    st = _debt_state
    st['last_sigs_t'] = trig_sigs
    st['last_sigs_d'] = data_sigs
    st['last_rows_t'] = cur_rows[:]
    st['year_label'] = year_label
    st['last_check'] = now

    lines: List[str] = []
    trig_done = False

    # ---------- 1. Триггерная проверка (пн, до зачёта) ----------
    if (wd == 0 or force_weekday) and st['trig_committed_on'] != today:
        if st['trig_baseline'] is None or len(st['trig_baseline']) != n:
            # Первая проверка — фиксируем базовые отпечатки столбца
            st['trig_baseline'] = trig_sigs[:]
            st['trig_rows_baseline'] = cur_rows[:]
        elif all(trig_sigs[i] != st['trig_baseline'][i] for i in range(n)):
            # Столбец «Количество» изменился в ОБОИХ таблицах — недельный зачёт
            _commit_tables(tmps)
            st['trig_baseline'] = trig_sigs[:]
            st['data_baseline'] = data_sigs[:]
            st['trig_rows_baseline'] = cur_rows[:]
            st['trig_committed_on'] = today
            trig_done = True
            rows = _load_debt_rows()
            count = len(rows) if rows is not None else "?"
            logger.info(f"Недельное обновление сверки зачтено ({count} записей)")
            await notify_debt_updated()
            gain = _rows_gain_text(st, cur_rows) or ""
            lines.append(f"✅ Сверка обновлена и разослана всем пользователям ({count} записей). {gain}")
        else:
            changed = [i + 1 for i in range(n) if trig_sigs[i] != st['trig_baseline'][i]]
            waiting = [i + 1 for i in range(n) if trig_sigs[i] == st['trig_baseline'][i]]
            if changed:
                lines.append(f"⏳ Столбец «{TRIG_COLUMN_HEADER}» изменился в таблице {changed}; "
                             f"жду таблицу {waiting}. Рассылка будет после обновления обеих таблиц.")

    # ---------- 2. Сверка-проверка (пн–пт, тихо) ----------
    if not trig_done:
        if wd in (5, 6) and not force_weekday:
            if any(not os.path.exists(p) for p in _debt_paths()):
                _commit_tables(tmps)
                rows = _load_debt_rows()
                count = len(rows) if rows is not None else "?"
                lines.append(f"♻️ Кэш-файлы отсутствовали — восстановлены ({count} записей).")
            else:
                _cleanup_files(tmps)
        elif st['data_baseline'] is None or len(st['data_baseline']) != n:
            _commit_tables(tmps)
            st['data_baseline'] = data_sigs[:]
            if st['trig_baseline'] is None:
                st['trig_baseline'] = trig_sigs[:]
                st['trig_rows_baseline'] = cur_rows[:]
            rows = _load_debt_rows()
            count = len(rows) if rows is not None else "?"
            lines.append(f"🔄 Лист «{DEBT_DATA_SHEET}» зафиксирован ({count} записей).")
        elif any(data_sigs[i] != st['data_baseline'][i] for i in range(n)):
            changed = [i + 1 for i in range(n) if data_sigs[i] != st['data_baseline'][i]]
            _commit_tables(tmps)
            st['data_baseline'] = data_sigs[:]
            rows = _load_debt_rows()
            count = len(rows) if rows is not None else "?"
            gain = _rows_gain_text(st, cur_rows)
            extra = f" {gain}" if gain else ""
            lines.append(f"🔄 Лист «{DEBT_DATA_SHEET}» изменился в таблице(ах) {changed} — "
                         f"активная сверка обновлена ({count} записей), без рассылки.{extra}")
        else:
            if any(not os.path.exists(p) for p in _debt_paths()):
                _commit_tables(tmps)
                rows = _load_debt_rows()
                count = len(rows) if rows is not None else "?"
                lines.append(f"♻️ Кэш-файлы отсутствовали — восстановлены ({count} записей).")
            else:
                _cleanup_files(tmps)
                if not lines:
                    lines.append("— Изменений в таблицах нет.")

    _save_debt_state()
    return "\n".join(lines) if lines else "— Изменений нет."


async def debt_refresh_loop():
    if not DEBT_SHEET_IDS:
        logger.warning("DEBT_SHEET_IDS не задан — обновление сверки отключено")
        return

    try:
        status = await startup_check()
        logger.info(f"Стартовая загрузка сверки: {status}")
    except Exception as e:
        logger.error(f"Ошибка стартовой загрузки сверки: {e}")

    while True:
        now = datetime.now(MSK)
        wd = now.weekday()

        # YEAR: недельная проверка метки года
        if (wd == DEBT_YEAR_CHECK_WD and now.hour >= DEBT_YEAR_CHECK_HOUR
                and DEBT_SHEET_IDS):
            try:
                paths = _debt_paths()
                if any(os.path.exists(p) for p in paths):
                    year_val = await asyncio.to_thread(
                        _read_year_cell, paths[0], DEBT_YEAR_CELL)
                    _debt_state['year_label'] = year_val
                    _save_debt_state()
                    cur_year = str(now.year)
                    if year_val is None or cur_year not in (year_val or ""):
                        if _debt_state.get('year_warned_for') != cur_year:
                            _debt_state['year_warned_for'] = cur_year
                            _save_debt_state()
                            users = load_users()
                            a = users.get(str(ADMIN_ID))
                            if a and a.get('chat_id'):
                                await bot.send_message(
                                    chat_id=a['chat_id'],
                                    text=("⚠️ **Проверьте ссылку на таблицы сверки!**\n\n"
                                          f"Метка года в {DEBT_YEAR_CELL}: "
                                          f"«{year_val if year_val else '—'}», ожидается {cur_year}.\n\n"
                                          "Если наступил новый год — обновите DEBT_SHEET_IDS в .env "
                                          "и перезапустите бота. До этого сверка берётся из прежних таблиц."),
                                    format="markdown")
                            logger.warning(f"Метка года сверки неактуальна: {year_val!r}")
            except Exception as e:
                logger.error(f"Ошибка проверки метки года: {e}")

        if wd in (5, 6):
            await asyncio.sleep(3600)
            continue

        interval_min = DEBT_CHECK_MONDAY_MINUTES if wd == 0 else DEBT_CHECK_WEEKDAY_MINUTES
        await asyncio.sleep(interval_min * 60)

        try:
            status = await scheduled_check()
            logger.info(f"Плановая проверка сверки: {status}")
        except Exception as e:
            logger.error(f"Ошибка плановой проверки сверки: {e}")


# ---------- ОБРАБОТКА USERINFO ----------

def _esc_html(text) -> str:
    """Экранирует спецсимволы HTML (&, <, >) для безопасного вывода в HTML-режиме."""
    if text is None:
        return ""
    s = str(text)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def format_user_info_from_dict(user_id: int, user_data: Dict) -> str:
    """Карточка контакта в HTML-формате (MAX поддерживает упоминания только так).
    Имя становится кликабельной ссылкой на чат с пользователем —
    клик открывает чат с этим пользователем."""
    fw_chat_id = user_data.get('chat_id', '—')
    fw_name = user_data.get('name')

    # Имя — кликабельная HTML-ссылка на чат с пользователем (если есть user_id)
    if fw_name:
        try:
            uid_int = int(user_id)
            if uid_int > 0:
                name_field = f'<a href="max://user/{uid_int}">{_esc_html(fw_name)}</a>'
            else:
                name_field = _esc_html(fw_name)
        except (TypeError, ValueError):
            name_field = _esc_html(fw_name)
    else:
        name_field = '—'

    return (
        "👤 <b>Данные контакта:</b>\n\n"
        f"• <b>User ID:</b> <code>{_esc_html(user_id)}</code>\n"
        f"• <b>Chat ID:</b> <code>{_esc_html(fw_chat_id)}</code>\n"
        f"• <b>Имя:</b> {name_field}\n"
    )


def format_forwarded_user_info(sender) -> str:
    """Карточка контакта из пересланного сообщения (HTML-формат).
    Имя — кликабельная ссылка на чат с пользователем при наличии fw_user_id."""
    fw_user_id = getattr(sender, 'user_id', None)
    fw_chat_id = getattr(sender, 'chat_id', None)
    fw_first_name = getattr(sender, 'first_name', None)
    fw_last_name = getattr(sender, 'last_name', None)

    if not fw_chat_id and fw_user_id:
        users = load_users()
        udata = users.get(str(fw_user_id))
        if udata:
            fw_chat_id = udata.get('chat_id')

    fio = " ".join(filter(None, [fw_last_name, fw_first_name])) or None

    # Имя — кликабельная HTML-ссылка на чат с пользователем
    if fio:
        try:
            if fw_user_id:
                uid_int = int(fw_user_id)
                if uid_int > 0:
                    name_field = f'<a href="max://user/{uid_int}">{_esc_html(fio)}</a>'
                else:
                    name_field = _esc_html(fio)
            else:
                name_field = _esc_html(fio)
        except (TypeError, ValueError):
            name_field = _esc_html(fio)
    else:
        name_field = '—'

    return (
        "👤 <b>Данные контакта:</b>\n\n"
        f"• <b>User ID:</b> <code>{_esc_html(fw_user_id) if fw_user_id else '—'}</code>\n"
        f"• <b>Chat ID:</b> <code>{_esc_html(fw_chat_id) if fw_chat_id else '—'}</code>\n"
        f"• <b>Имя:</b> {name_field}\n"
    )


def _find_user_by_any_id(uid: int) -> Optional[tuple]:
    """Ищет контакт в базе по User ID (ключ) или Chat ID (поле chat_id).
    Возвращает (user_id, данные) или None, если контакт не найден."""
    users = load_users()
    data = users.get(str(uid))
    if data:
        return uid, data
    for str_uid, udata in users.items():
        chat_id_u = udata.get('chat_id')
        if chat_id_u is None:
            continue
        try:
            if int(chat_id_u) == uid:
                try:
                    return int(str_uid), udata
                except ValueError:
                    continue
        except (TypeError, ValueError):
            continue
    return None


async def process_userinfo_input(event, user_id: int, chat_id: int, text: str):
    link = getattr(event.message, 'link', None)
    if link and getattr(link, 'sender', None):
        sender = link.sender
        logger.info(f"Пересылка обнаружена, отправитель: user_id={getattr(sender, 'user_id', None)}")
        text_out = format_forwarded_user_info(sender)
        await bot.send_message(
            chat_id=chat_id,
            text=text_out,
            format="html",
            attachments=[create_main_menu(user_id)]
        )
        user_states.pop(user_id, None)
        return

    # Админ ввёл User ID или Chat ID цифрами — ищем контакт в базе
    stripped = (text or "").strip().strip('`').lstrip('+')
    if stripped.isdigit():
        found = _find_user_by_any_id(int(stripped))
        if found:
            f_uid, f_data = found
            logger.info(f"UserInfo по ID: запрошен {stripped}, найден user_id={f_uid} (запросил админ {user_id})")
            await bot.send_message(
                chat_id=chat_id,
                text=format_user_info_from_dict(f_uid, f_data),
                format="html",
                attachments=[create_main_menu(user_id)]
            )
            user_states.pop(user_id, None)
            return

        logger.info(f"UserInfo по ID: контакт {stripped} не найден (запросил админ {user_id})")
        await bot.send_message(
            chat_id=chat_id,
            text=("⛔ **Контакт не найден.**\n\n"
                  f"В базе бота нет контакта с User ID / Chat ID `{stripped}`.\n"
                  "Карточка доступна, только если контакт хотя бы раз заходил в бота.\n\n"
                  "Проверьте ID и отправьте его снова — либо нажмите «Отмена»."),
            format="markdown",
            attachments=[create_cancel_menu()]
        )
        # Состояние не сбрасываем — админ может ввести другой ID или отменить
        return

    users = load_users()
    user_data = users.get(str(user_id))
    if user_data:
        text_out = format_user_info_from_dict(user_id, user_data)
    else:
        text_out = (
            "👤 **Ваши данные:**\n\n"
            f"• **User ID:** `{user_id}`\n"
            f"• **Chat ID:** `{chat_id}`\n"
            f"• В базе данных о вас пока нет записи."
        )
    await bot.send_message(
        chat_id=chat_id,
        text=text_out,
        format="markdown",
        attachments=[create_main_menu(user_id)]
    )
    user_states.pop(user_id, None)


# ---------- ОБРАБОТЧИКИ ----------

@dp.bot_started()
async def handle_started(event: BotStarted):
    user = event.user
    user_id = user.user_id
    chat_id = event.chat_id
    name = getattr(user, 'first_name', 'друг')
    username = getattr(user, 'username', None)
    ensure_user(user_id, chat_id, username, name)
    await bot.send_message(
        chat_id=chat_id,
        text=f"👋 Привет, {name}!\n\nВыберите действие:",
        attachments=[create_main_menu(user_id)]
    )


@dp.message_created(Command('start'))
async def cmd_start(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    name = getattr(user, 'first_name', 'друг')
    username = getattr(user, 'username', None)
    ensure_user(user_id, chat_id, username, name)
    user_states.pop(user_id, None)
    user_clarify_data.pop(user_id, None)
    admin_debt_data.pop(user_id, None)
    await bot.send_message(
        chat_id=chat_id,
        text=f"👋 Привет, {name}!\n\nВыберите действие:",
        attachments=[create_main_menu(user_id)]
    )


# MENU: всегда доступное меню по /menu (синоним /start без приветствия)
@dp.message_created(Command('menu'))
async def cmd_menu(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    user_states.pop(user_id, None)
    user_clarify_data.pop(user_id, None)
    admin_debt_data.pop(user_id, None)
    await bot.send_message(
        chat_id=chat_id,
        text="🏠 Главное меню:",
        attachments=[create_main_menu(user_id)]
    )


async def send_help(chat_id: int, user_id: int):
    admin = is_admin(user_id)
    text = "📋 **Справка**\n\n"
    text += "Кнопки меню доступны всегда — отправьте /menu или /start.\n\n"
    text += "Кнопки:\n"
    text += "• 📋 Условия – условия по регионам\n"
    if admin:
        text += "• 💳 Долг – проверить задолженность по любому ФИО\n"
    else:
        text += "• 💳 Долг – проверить задолженность\n"
    text += "• ℹ️ О боте – информация о боте\n"
    if admin:
        text += "• 📢 Новости – сделать рассылку (только для админа)\n"
        text += "• 👤 ID контакта – получить данные пользователя (только для админа)\n"
        text += "• 🔄 Обновить сверку / 📊 Статус сверки – кнопки (только для админа)\n"
    if is_owner(user_id):
        text += "• 👥 Админы – список админов (владелец)\n"
        text += "• ➕/➖ Админы – назначить/снять админа (владелец)\n"

    text += "\n**Доступные команды:**\n"
    text += "• /menu — Главное меню (кнопки всегда под рукой)\n"
    text += "• /start — Главное меню\n"
    text += "• /help — Эта справка\n"
    text += "• /debt — Узнать долг\n"
    text += "• /change_name — Сменить ФИО\n"
    text += "• /cancel — Отменить текущее действие\n"
    if admin:
        text += "• /broadcast — Начать рассылку (только админ)\n"
        text += "• /sentlog — журнал отправок бота (только админ)\n"
        text += "• /userinfo — ID контакта (только админ)\n• /sentlog — журнал отправок бота (только админ)\n"
        text += "• /refresh_debt — Проверить таблицы сверки сейчас (только админ)\n"
        text += "• /debt_status — Статус сверки (только админ)\n"
        text += "• /adminlist — Список админов\n"
    if is_owner(user_id):
        text += "• /addadmin <id> — Назначить админа (владелец)\n"
        text += "• /deladmin <id> — Снять админа (владелец)\n"

    await bot.send_message(
        chat_id=chat_id,
        text=text,
        format="markdown",
        attachments=[create_main_menu(user_id)]
    )


@dp.message_created(Command('help'))
async def cmd_help(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await send_help(chat_id, user_id)


@dp.message_callback(F.callback.payload == "help_command")
async def callback_help_command(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="❓ Справка")
    await send_help(chat_id, user_id)


@dp.message_created(Command('change_name'))
async def cmd_change_name(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await start_change_name(user_id, chat_id, event)


async def start_change_name(user_id: int, chat_id: int, event):
    clear_user_full_name(user_id)
    user_states.pop(user_id, None)
    user_clarify_data.pop(user_id, None)
    admin_debt_data.pop(user_id, None)
    user_states[user_id] = UserState.DEBT_NAME_MODE
    await bot.send_message(
        chat_id=chat_id,
        text="🔄 **Смена ФИО**\n\n"
             "Введите ваши ФИО, на которые оформлена основная доверенность.\n"
             "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )
    if hasattr(event, 'answer'):
        await event.answer(notification="🔄 Введите новые ФИО")


@dp.message_created(Command('debt'))
async def cmd_debt(event: MessageCreated):
    await handle_debt_request(event.message.sender.user_id, event.message.recipient.chat_id, event)


@dp.message_callback(F.callback.payload == "debt")
async def callback_debt(event: MessageCallback):
    user_id = event.callback.user.user_id
    chat_id = event.message.recipient.chat_id
    await handle_debt_request(user_id, chat_id, event)


async def handle_debt_request(user_id: int, chat_id: int, event):
    ensure_user(user_id, chat_id, None, None)

    if is_admin(user_id):
        user_states[user_id] = UserState.ADMIN_DEBT_MODE
        user_clarify_data.pop(user_id, None)
        admin_debt_data.pop(user_id, None)
        await bot.send_message(
            chat_id=chat_id,
            text="💳 **Проверка долга клиента**\n\n"
                 "Введите ФИО контакта, для которого нужно узнать задолженность.\n"
                 "Например: *Иванов Иван* или *Иванов Иван Иванович*\n\n"
                 "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
            format="markdown",
            attachments=[create_cancel_menu()]
        )
        if hasattr(event, 'answer'):
            await event.answer(notification="💳 Введите ФИО клиента")
        return

    users = load_users()
    user_data = users.get(str(user_id))
    full_name = user_data.get('full_name') if user_data else None

    if full_name:
        await process_debt_search(user_id, chat_id, full_name)
    else:
        user_states[user_id] = UserState.DEBT_NAME_MODE
        user_clarify_data.pop(user_id, None)
        await bot.send_message(
            chat_id=chat_id,
            text="💳 **Проверка долга**\n\n"
                 "Укажите ваши ФИО, на которые оформлена основная доверенность.\n"
                 "Например: *Иванов Иван* или *Иванов Иван Иванович*\n\n"
                 "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
            format="markdown",
            attachments=[create_cancel_menu()]
        )
        if hasattr(event, 'answer'):
            await event.answer(notification="💳 Введите ФИО")


async def process_debt_search(user_id: int, chat_id: int, full_name: str):
    matches = await search_debt_with_recovery(full_name)

    if matches is None:
        user_states.pop(user_id, None)
        await bot.send_message(
            chat_id=chat_id,
            text="⚠️ **Сервис проверки долга временно недоступен.**\n\n"
                 "Пожалуйста, попробуйте позже или обратитесь к менеджеру.",
            attachments=[create_main_menu(user_id)]
        )
        return

    if not matches:
        await bot.send_message(
            chat_id=chat_id,
            text="❌ **Доверенность с данными ФИО не найдена.**\n\n"
                 "Пожалуйста, проверьте правильность ввода или обратитесь к менеджеру.\n"
                 "Вы можете сменить ФИО с помощью кнопки ниже.",
            format="markdown",
            attachments=[create_debt_result_menu()]
        )
        return

    groups = group_matches(matches)

    if len(groups) == 1:
        _, items = next(iter(groups.items()))
        display = clean_name(items[0]['name'])
        set_user_full_name(user_id, display)
        await send_debt_result(chat_id, user_id, matches)
        return

    base = ' '.join(_norm_word(w) for w in full_name.strip().split()[:2])
    variants = {key: clean_name(items[0]['name']) for key, items in groups.items()}

    user_clarify_data[user_id] = {'base': base, 'variants': variants}
    user_states[user_id] = UserState.DEBT_CLARIFY_MODE

    variants_text = "\n".join([f"• {name}" for name in variants.values()])
    await bot.send_message(
        chat_id=chat_id,
        text=f"🔍 **Найдено несколько вариантов:**\n\n{variants_text}\n\n"
             "Уточните, пожалуйста, ваше **отчество** (или введите полное ФИО).\n"
             "Например: *Иванович* или *Иванов Иван Иванович*\n\n"
             "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


async def process_admin_debt_search(user_id: int, chat_id: int, full_name: str):
    matches = await search_debt_with_recovery(full_name)

    if matches is None:
        await bot.send_message(
            chat_id=chat_id,
            text="⚠️ **Сервис проверки долга временно недоступен.**\n\nПопробуйте позже.",
            attachments=[create_main_menu(user_id)]
        )
        return

    if not matches:
        await bot.send_message(
            chat_id=chat_id,
            text=f"❌ **Доверенность с ФИО «{full_name}» не найдена.**\n\n"
                 "Проверьте ввод и попробуйте ещё раз.",
            format="markdown",
            attachments=[create_admin_debt_menu()]
        )
        return

    groups = group_matches(matches)

    if len(groups) == 1:
        await send_debt_result(chat_id, user_id, matches, menu=create_admin_debt_menu())
        return

    base = ' '.join(_norm_word(w) for w in full_name.strip().split()[:2])
    variants = {key: clean_name(items[0]['name']) for key, items in groups.items()}
    admin_debt_data[user_id] = {'base': base, 'variants': variants}
    user_states[user_id] = UserState.ADMIN_DEBT_CLARIFY_MODE

    variants_text = "\n".join([f"• {name}" for name in variants.values()])
    await bot.send_message(
        chat_id=chat_id,
        text=f"🔍 **Найдено несколько вариантов:**\n\n{variants_text}\n\n"
             "Уточните **отчество** клиента (или введите полное ФИО).\n"
             "Например: *Иванович* или *Иванов Иван Иванович*\n\n"
             "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


@dp.message_callback(F.callback.payload == "change_name")
async def callback_change_name(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await start_change_name(user_id, chat_id, event)


@dp.message_callback(F.callback.payload == "admin_debt_again")
async def callback_admin_debt_again(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id

    if not is_admin(user_id):
        await event.answer(notification="⛔ Доступно только администратору")
        return

    await handle_debt_request(user_id, chat_id, event)


# ---------- КОМАНДА /USERINFO (ТОЛЬКО АДМИН) ----------

@dp.message_created(Command('userinfo'))
async def cmd_userinfo(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))

    if not is_admin(user_id):
        await event.message.answer("⛔ Эта команда доступна только администратору.")
        return

    user_states[user_id] = UserState.WAITING_FOR_FORWARD
    await event.message.answer(
        "👤 **Получение информации о контакте**\n\n"
        "• Перешлите мне сообщение от другого контакта — покажу его данные.\n"
        "• Введите User ID или Chat ID цифрами — найду контакт в базе.\n"
        "• Напишите любое сообщение — покажу ваши данные.\n\n"
        "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


@dp.message_created(Command('sentlog'))
async def cmd_sentlog(event: MessageCreated):
    """Журнал отправок бота: кому и какие массовые сообщения отправлены (только админ)."""
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))

    if not is_admin(user_id):
        await event.message.answer("⛔ Эта команда доступна только администратору.")
        return

    body = getattr(event.message, 'body', None)
    arg = ''
    if body and getattr(body, 'text', None):
        parts = body.text.split()
        if len(parts) > 1:
            arg = parts[1].lower()

    # /sentlog clear — очистить журнал
    if arg in ('clear', 'очистить'):
        entries = _load_sent_log()
        try:
            os.makedirs('data', exist_ok=True)
            with open(SENT_LOG_FILE, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False)
            await event.message.answer(
                f"🗑 Журнал отправок очищен (удалено записей: {len(entries)}).",
                attachments=[create_main_menu(user_id)]
            )
        except Exception as e:
            logger.error(f"Не удалось очистить журнал отправок: {e}")
            await event.message.answer("❌ Не удалось очистить журнал (см. лог).")
        return

    n = 10
    if arg.isdigit():
        n = max(1, min(int(arg), 50))

    entries = _load_sent_log()
    if not entries:
        await event.message.answer(
            "📭 Журнал отправок пуст.\n\n"
            "Сюда попадают массовые отправки бота: рассылки новостей и уведомления о сверке.",
            attachments=[create_main_menu(user_id)]
        )
        return

    total = len(entries)
    shown = entries[-n:]
    lines = [f"📨 Журнал отправок — последние {len(shown)} из {total}\n"]
    for e in shown:
        snippet = (e.get('text') or '').replace('\n', ' ')
        if len(snippet) > 60:
            snippet = snippet[:60] + '…'
        lines.append(
            f"{e.get('ts', '—')} | {e.get('source', '—')}\n"
            f"   → {e.get('name', '—')} (ID {e.get('user_id', '—')}): «{snippet}»"
        )
    lines.append("\nПолные записи — в файле data/sent_log.json на сервере. "
                 "Очистить журнал: /sentlog clear")
    await event.message.answer("\n".join(lines), attachments=[create_main_menu(user_id)])


@dp.message_callback(F.callback.payload == "userinfo")
async def callback_userinfo(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))

    if not is_admin(user_id):
        await event.answer(notification="⛔ Доступно только администратору")
        await bot.send_message(
            chat_id=chat_id,
            text="⛔ Эта функция доступна только администратору.",
            attachments=[create_main_menu(user_id)]
        )
        return

    user_states[user_id] = UserState.WAITING_FOR_FORWARD
    await event.answer(notification="👤 Ожидаю ввод")
    await bot.send_message(
        chat_id=chat_id,
        text="👤 **Получение информации о контакте**\n\n"
             "• Перешлите мне сообщение от другого контакта — покажу его данные.\n"
             "• Введите User ID или Chat ID цифрами — найду контакт в базе.\n"
             "• Напишите любое сообщение — покажу ваши данные.\n\n"
             "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


# ---------- УПРАВЛЕНИЕ АДМИНАМИ ----------

@dp.message_created(Command('addadmin'))
async def cmd_addadmin(event: MessageCreated):
    user_id = event.message.sender.user_id
    chat_id = event.message.recipient.chat_id

    if not is_owner(user_id):
        await event.message.answer("⛔ Управлять админами может только владелец бота.")
        return

    body = getattr(event.message, 'body', None)
    text = (getattr(body, 'text', None) or '') if body else ''
    m = re.match(r'^/addadmin\s+(\d+)\s*$', text.strip())
    if not m:
        user_states[user_id] = UserState.WAITING_ADDADMIN_ID
        await event.message.answer(
            "➕ **Назначение админа**\n\n"
            "Отправьте user_id пользователя (только число).\n"
            "Он должен хотя бы раз писать боту.\n"
            "Для отмены — /cancel.",
            format="markdown",
            attachments=[create_cancel_menu()])
        return

    new_id = int(m.group(1))
    if new_id == ADMIN_ID:
        await event.message.answer("ℹ️ Владелец уже является администратором.")
        return
    if new_id in _dynamic_admins:
        await event.message.answer(f"ℹ️ Пользователь {new_id} уже администратор.")
        return

    if str(new_id) not in load_users():
        await event.message.answer(
            f"⚠️ Пользователь {new_id} не найден в базе.\n"
            "Он должен хотя бы раз написать боту (/start), прежде чем его можно будет назначить.")
        return

    _dynamic_admins.add(new_id)
    _save_admins()
    logger.info(f"Владелец назначил админа: {new_id}")
    await event.message.answer(f"✅ Пользователь `{new_id}` назначен администратором.", format="markdown")
    try:
        await bot.send_message(chat_id=new_id, text="🎉 Вам выданы права администратора бота.")
    except Exception:
        pass


@dp.message_created(Command('deladmin'))
async def cmd_deladmin(event: MessageCreated):
    user_id = event.message.sender.user_id
    chat_id = event.message.recipient.chat_id

    if not is_owner(user_id):
        await event.message.answer("⛔ Управлять админами может только владелец бота.")
        return

    body = getattr(event.message, 'body', None)
    text = (getattr(body, 'text', None) or '') if body else ''
    m = re.match(r'^/deladmin\s+(\d+)\s*$', text.strip())
    if not m:
        user_states[user_id] = UserState.WAITING_DELADMIN_ID
        await event.message.answer(
            "➖ **Снятие админа**\n\n"
            "Отправьте user_id администратора (только число).\n"
            "Для отмены — /cancel.",
            format="markdown",
            attachments=[create_cancel_menu()])
        return

    del_id = int(m.group(1))
    if del_id == ADMIN_ID:
        await event.message.answer("⛔ Нельзя снять самого владельца.")
        return
    if del_id not in _dynamic_admins:
        await event.message.answer(f"ℹ️ Пользователь {del_id} не является динамическим админом.")
        return

    _dynamic_admins.discard(del_id)
    _save_admins()
    logger.info(f"Владелец снял админа: {del_id}")
    await event.message.answer(f"✅ Права администратора сняты с `{del_id}`.", format="markdown")
    try:
        await bot.send_message(chat_id=del_id, text="ℹ️ Права администратора бота отозваны.")
    except Exception:
        pass


@dp.message_created(Command('adminlist'))
async def cmd_adminlist(event: MessageCreated):
    user_id = event.message.sender.user_id

    if not is_admin(user_id):
        await event.message.answer("⛔ Только для администраторов.")
        return

    lines = ["👥 **Администраторы**\n"]
    lines.append(f"• 👑 {ADMIN_ID} — владелец")
    for did in sorted(_dynamic_admins):
        lines.append(f"• {did} — назначен")
    if not _dynamic_admins:
        lines.append("_(назначенных админов нет)_")
    await event.message.answer("\n".join(lines), format="markdown")


@dp.message_callback(F.callback.payload == "adminlist")
async def callback_adminlist(event: MessageCallback):
    user_id = event.callback.user.user_id
    chat_id = event.message.recipient.chat_id

    if not is_admin(user_id):
        await event.answer(notification="⛔ Только администратору")
        return

    await event.answer(notification="👥 Админы")
    lines = ["👥 **Администраторы**\n"]
    lines.append(f"• 👑 {ADMIN_ID} — владелец")
    for did in sorted(_dynamic_admins):
        lines.append(f"• {did} — назначен")
    if not _dynamic_admins:
        lines.append("_(назначенных админов нет)_")
    await bot.send_message(chat_id=chat_id, text="\n".join(lines), format="markdown")


# BTN: кнопки ➕/➖ Админы (диалог ввода ID)
@dp.message_callback(F.callback.payload == "addadmin")
async def callback_addadmin(event: MessageCallback):
    user_id = event.callback.user.user_id
    chat_id = event.message.recipient.chat_id

    if not is_owner(user_id):
        await event.answer(notification="⛔ Только владельцу")
        return

    await event.answer(notification="Введите ID")
    user_states[user_id] = UserState.WAITING_ADDADMIN_ID
    await bot.send_message(
        chat_id=chat_id,
        text="➕ **Назначение админа**\n\n"
             "Отправьте user_id пользователя (только число).\n"
             "Он должен хотя бы раз писать боту.\n"
             "Для отмены — /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


@dp.message_callback(F.callback.payload == "deladmin")
async def callback_deladmin(event: MessageCallback):
    user_id = event.callback.user.user_id
    chat_id = event.message.recipient.chat_id

    if not is_owner(user_id):
        await event.answer(notification="⛔ Только владельцу")
        return

    await event.answer(notification="Введите ID")
    user_states[user_id] = UserState.WAITING_DELADMIN_ID
    await bot.send_message(
        chat_id=chat_id,
        text="➖ **Снятие админа**\n\n"
             "Отправьте user_id администратора (только число).\n"
             "Для отмены — /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


# ---------- АДМИН-КОМАНДЫ СВЕРКИ ----------

def _debt_status_text() -> str:
    now = datetime.now(MSK)
    wd_names = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]

    lines = ["📊 **Статус сверки**\n"]
    lines.append(f"• Сейчас: {wd_names[now.weekday()]}, {now:%H:%M} МСК")
    if now.weekday() in (5, 6):
        lines.append("• Сегодня выходной — плановых проверок нет")
    elif now.weekday() == 0:
        lines.append(f"• Расписание: триггеры + сверка каждые {DEBT_CHECK_MONDAY_MINUTES} мин")
    else:
        lines.append(f"• Расписание: сверка каждые {DEBT_CHECK_WEEKDAY_MINUTES} мин")
    lines.append(f"• Лист с долгами: «{DEBT_DATA_SHEET}»")
    lines.append(f"• Лист-триггер: «{DEBT_SOURCE_SHEET}»")

    st = _debt_state
    # STATUS-YEAR: метка года значением
    yl = st.get('year_label')
    lines.append(f"• Метка года: {yl if yl else '—'}")

    today = now.date().isoformat()
    if st['trig_committed_on'] == today:
        lines.append("• Триггеры: ✅ недельное обновление сегодня зачтено — до следующего пн не проверяются\n")
    else:
        lines.append("• Триггеры: ожидание обновления (проверяются только в пн)\n")

    for i, sid in enumerate(DEBT_SHEET_IDS, 1):
        parts = [f"**Таблица {i}**"]
        parts.append(f"• ID: `{sid[:12]}…`")
        tb = st['trig_baseline']
        lt = st['last_sigs_t']
        if tb and lt and i - 1 < len(lt) and lt[i - 1]:
            if tb[i - 1] != lt[i - 1]:
                parts.append("• Триггер (ст. «Количество»): 🟡 изменился (ждём остальные)")
            else:
                parts.append("• Триггер (ст. «Количество»): ✅ без изменений")
        trb = st.get('trig_rows_baseline')
        lrt = st.get('last_rows_t')
        if trb and lrt and i - 1 < len(trb) and i - 1 < len(lrt) \
                and trb[i - 1] is not None and lrt[i - 1] is not None:
            if lrt[i - 1] != trb[i - 1]:
                parts.append(f"• Выгрузка: 🟡 строк {lrt[i - 1]} (baseline {trb[i - 1]})")
            else:
                parts.append(f"• Выгрузка: ✅ {lrt[i - 1]} строк")
        db = st['data_baseline']
        ld = st['last_sigs_d']
        if db and ld and i - 1 < len(ld) and ld[i - 1]:
            if db[i - 1] != ld[i - 1]:
                parts.append("• Сверка: 🟡 в Google новее активной (будет подхвачена)")
            else:
                parts.append("• Сверка: ✅ актуальна")
        lines.append("\n".join(parts))

    if st['last_check']:
        lines.append(f"\n• Последняя проверка: {st['last_check']:%d.%m.%Y %H:%M:%S} МСК")

    return "\n".join(lines)


@dp.message_created(Command('refresh_debt'))
async def cmd_refresh_debt(event: MessageCreated):
    user_id = event.message.sender.user_id
    chat_id = event.message.recipient.chat_id

    if not is_admin(user_id):
        await event.message.answer("⛔ Эта команда только для администратора.")
        return

    if not DEBT_SHEET_IDS:
        await event.message.answer("⚠️ DEBT_SHEET_IDS не задан в настройках.")
        return

    await event.message.answer("🔎 Проверяю таблицы…")
    status = await scheduled_check(force_weekday=True)
    await event.message.answer(status, format="markdown")


@dp.message_created(Command('debt_status'))
async def cmd_debt_status(event: MessageCreated):
    user_id = event.message.sender.user_id
    chat_id = event.message.recipient.chat_id

    if not is_admin(user_id):
        await event.message.answer("⛔ Эта команда только для администратора.")
        return

    if not DEBT_SHEET_IDS:
        await event.message.answer("❌ DEBT_SHEET_IDS не задан.")
        return

    await event.message.answer(_debt_status_text(), format="markdown")


# BTN: колбэки админ-кнопок сверки
@dp.message_callback(F.callback.payload == "refresh_debt")
async def callback_refresh_debt(event: MessageCallback):
    user_id = event.callback.user.user_id
    chat_id = event.message.recipient.chat_id

    if not is_admin(user_id):
        await event.answer(notification="⛔ Только администратору")
        return

    await event.answer(notification="🔎 Проверяю таблицы…")
    status = await scheduled_check(force_weekday=True)
    await bot.send_message(chat_id=chat_id, text=status, format="markdown")


@dp.message_callback(F.callback.payload == "debt_status")
async def callback_debt_status(event: MessageCallback):
    user_id = event.callback.user.user_id
    chat_id = event.message.recipient.chat_id

    if not is_admin(user_id):
        await event.answer(notification="⛔ Только администратору")
        return

    await event.answer(notification="📊 Статус сверки")
    await bot.send_message(chat_id=chat_id, text=_debt_status_text(), format="markdown")


# ---------- ОСТАЛЬНЫЕ ОБРАБОТЧИКИ ----------

@dp.message_callback(F.callback.payload == "info")
async def callback_info(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="ℹ️ О боте")
    await bot.send_message(
        chat_id=chat_id,
        text="ℹ️ **О боте**\n\nВерсия: 1.9\nНазначение: автоматизация новостных рассылок для подписчиков и предоставление общей информации по сотрудничеству.",
        format="markdown",
        attachments=[create_main_menu(user_id)]
    )


@dp.message_callback(F.callback.payload == "conditions")
async def callback_conditions(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="📋 Выбор региона")
    await bot.send_message(
        chat_id=chat_id,
        text="📋 **Выберите ваш регион**\n\nДля просмотра условий выберите нужный регион:",
        format="markdown",
        attachments=[create_regions_menu()]
    )


@dp.message_callback(F.callback.payload == "region_1")
async def callback_region_1(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="📍 Краснодар")
    text = (
        "📋 **Условия по Краснодару**\n\n"
        "1. Товар строго для кормления животных.\n"
        "2. Требуется ежедневный вывоз, либо по согласованному графику с директором магазина.\n"
        "3. Перед приездом нужно оповестить директора магазина.\n"
        "4. Гниль и плесень не забираем. Перед взвешиванием можно отложить гниль в сторону, определив по внешнему виду.\n"
        "5. Тара под списание должна быть своей (тара из магазина возвратная).\n"
        "6. Оплата производится по накладным (с понедельника по воскресенье включительно) в воскресенье/понедельник на р/счёт или по qr-коду.\n"
        "7. При жалобах с магазина, задержках оплаты и других нарушениях условий доверенность аннулируется.\n"
        "8. По вопросам или по сложным ситуациям можно сообщить по номеру: +7 953 641-69-93 Александр.\n\n"
        "Для оформления доверенности необходимы полные фамилия, имя, отчество водителя, серия и номер паспорта, кем и когда выдан, а также заявление, подтверждающее наличие животных.\n"
        "Заявление заполняется каждым водителем, который будет забирать списание. Заявление можно написать от руки на А4 по образцу, поставить подпись, и фото заполненного заявления прислать мне.\n"
        "Доверенность оформляется на человека, который будет забирать просрок.\n\n"
        "• Цена: 15 руб. за 1 кг продукции;\n"
        "• Ассортимент: Хлеб, хлебобулочная продукция, овощи, фрукты, молочная, мясная гастрономия и бакалея;\n"
    )
    await bot.send_message(chat_id=chat_id, text=text, format="markdown")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите нужный документ:**",
        format="markdown",
        attachments=[create_materials_menu("region1")]
    )


@dp.message_callback(F.callback.payload == "region_2")
async def callback_region_2(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="📍 КМВ, КЧР")
    text = (
        "📋 **Условия по Ставропольскому краю и КЧР**\n\n"
        "1. Товар строго для кормления животных.\n"
        "2. Требуется ежедневный вывоз, либо по согласованному графику с директором магазина.\n"
        "3. Перед приездом нужно оповестить директора магазина.\n"
        "4. Гниль и плесень не забираем. Перед взвешиванием можно отложить гниль в сторону, определив по внешнему виду.\n"
        "5. Тара под списание должна быть своей (тара из магазина возвратная).\n"
        "6. Оплата производится по накладным (с понедельника по воскресенье включительно) в воскресенье/понедельник на р/счёт или по qr-коду.\n"
        "7. При жалобах с магазина, задержках оплаты и других нарушениях условий доверенность аннулируется.\n"
        "8. По вопросам или по сложным ситуациям можно сообщить по номеру: +7 953 641-69-93 Александр.\n\n"
        "Для оформления доверенности необходимы полные фамилия, имя, отчество водителя, серия и номер паспорта, кем и когда выдан, а также заявление, подтверждающее наличие животных.\n"
        "Заявление заполняется каждым водителем, который будет забирать списание. Заявление можно написать от руки на А4 по образцу, поставить подпись, и фото заполненного заявления прислать мне.\n"
        "Доверенность оформляется на человека, который будет забирать просрок.\n\n"
        "• Цена: 17 руб. за 1 кг продукции;\n"
        "• Ассортимент: Хлеб, хлебобулочная продукция, овощи, фрукты, молочная, мясная гастрономия и бакалея;\n"
    )
    await bot.send_message(chat_id=chat_id, text=text, format="markdown")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите нужный документ:**",
        format="markdown",
        attachments=[create_materials_menu("region2")]
    )


@dp.message_callback(F.callback.payload == "region1_app")
async def callback_region1_app(event: MessageCallback):
    chat_id = event.message.recipient.chat_id
    await event.answer(notification="📄 Заявление")
    await send_image(chat_id, REGION1_APP, "📄 Заявление о наличии животных")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите другой документ:**",
        format="markdown",
        attachments=[create_materials_menu("region1")]
    )


@dp.message_callback(F.callback.payload == "region1_qr")
async def callback_region1_qr(event: MessageCallback):
    chat_id = event.message.recipient.chat_id
    await event.answer(notification="📱 QR-код")
    await send_image(chat_id, REGION1_QR, "📱 QR-код для оплаты")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите другой документ:**",
        format="markdown",
        attachments=[create_materials_menu("region1")]
    )


@dp.message_callback(F.callback.payload == "region1_bank")
async def callback_region1_bank(event: MessageCallback):
    chat_id = event.message.recipient.chat_id
    await event.answer(notification="💳 Реквизиты")
    await bot.send_message(chat_id=chat_id, text=f"**💳 Реквизиты для оплаты:**\n\n{REGION1_BANK_TEXT}", format="markdown")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите другой документ:**",
        format="markdown",
        attachments=[create_materials_menu("region1")]
    )


@dp.message_callback(F.callback.payload == "region2_app")
async def callback_region2_app(event: MessageCallback):
    chat_id = event.message.recipient.chat_id
    await event.answer(notification="📄 Заявление")
    await send_image(chat_id, REGION2_APP, "📄 Заявление о наличии животных")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите другой документ:**",
        format="markdown",
        attachments=[create_materials_menu("region2")]
    )


@dp.message_callback(F.callback.payload == "region2_qr")
async def callback_region2_qr(event: MessageCallback):
    chat_id = event.message.recipient.chat_id
    await event.answer(notification="📱 QR-код")
    await send_image(chat_id, REGION2_QR, "📱 QR-код для оплаты")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите другой документ:**",
        format="markdown",
        attachments=[create_materials_menu("region2")]
    )


@dp.message_callback(F.callback.payload == "region2_bank")
async def callback_region2_bank(event: MessageCallback):
    chat_id = event.message.recipient.chat_id
    await event.answer(notification="💳 Реквизиты")
    await bot.send_message(chat_id=chat_id, text=f"**💳 Реквизиты для оплаты:**\n\n{REGION2_BANK_TEXT}", format="markdown")
    await bot.send_message(
        chat_id=chat_id,
        text="📎 **Выберите другой документ:**",
        format="markdown",
        attachments=[create_materials_menu("region2")]
    )


@dp.message_callback(F.callback.payload == "back_to_regions")
async def callback_back_to_regions(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="🔙 Назад к регионам")
    await bot.send_message(
        chat_id=chat_id,
        text="📋 **Выберите регион:**",
        format="markdown",
        attachments=[create_regions_menu()]
    )


@dp.message_callback(F.callback.payload == "back_to_main")
async def callback_back_to_main(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    await event.answer(notification="🔙 Возврат")
    await bot.send_message(
        chat_id=chat_id,
        text="🏠 Главное меню:",
        attachments=[create_main_menu(user_id)]
    )


@dp.message_callback(F.callback.payload == "broadcast")
async def callback_broadcast(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))

    if not is_admin(user_id):
        await event.answer(notification="⛔ Доступно только администратору")
        await bot.send_message(
            chat_id=chat_id,
            text="⛔ Эта функция доступна только администратору.",
            attachments=[create_main_menu(user_id)]
        )
        return

    user_states[user_id] = UserState.BROADCAST_MODE
    await event.answer(notification="📢 Режим рассылки")
    await bot.send_message(
        chat_id=chat_id,
        text="📢 **Рассылка новостей**\n\n"
             "Введите текст, который нужно разослать всем пользователям.\n"
             "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


@dp.message_callback(F.callback.payload == "cancel_action")
async def callback_cancel_action(event: MessageCallback):
    user = event.callback.user
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    user_states.pop(user_id, None)
    user_clarify_data.pop(user_id, None)
    admin_debt_data.pop(user_id, None)
    await event.answer(notification="❌ Действие отменено")
    await bot.send_message(
        chat_id=chat_id,
        text="❌ Вы отменили текущее действие.",
        attachments=[create_main_menu(user_id)]
    )


@dp.message_created(Command('broadcast'))
async def cmd_broadcast(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))

    if not is_admin(user_id):
        await event.message.answer("⛔ Эта команда только для администратора.")
        return

    user_states[user_id] = UserState.BROADCAST_MODE
    await event.message.answer(
        "📢 **Рассылка новостей**\n\n"
        "Введите текст, который нужно разослать всем пользователям.\n"
        "Для отмены нажмите кнопку «Отмена» или введите команду /cancel.",
        format="markdown",
        attachments=[create_cancel_menu()]
    )


@dp.message_created(Command('cancel'))
async def cmd_cancel(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    ensure_user(user_id, chat_id, getattr(user, 'username', None), getattr(user, 'first_name', None))
    user_states.pop(user_id, None)
    user_clarify_data.pop(user_id, None)
    admin_debt_data.pop(user_id, None)
    await event.message.answer(
        "❌ Действие отменено.",
        attachments=[create_main_menu(user_id)]
    )


# ---------- ОТПРАВКА КАРТИНОК ----------

def _pick_upload_type():
    if UploadType is not None:
        for cand in ("image", "images"):
            for member in UploadType:
                if member.value == cand:
                    return member
    return "image"


def _extract_token(raw):
    if raw is None:
        return None
    if not isinstance(raw, str):
        raw = getattr(raw, "token", None) or getattr(raw, "value", None) or str(raw)
    s = str(raw).strip()
    if s.startswith("{"):
        try:
            j = json.loads(s)
            photos = j.get("photos") if isinstance(j, dict) else None
            if isinstance(photos, dict) and photos:
                # Эталон maxapi (_extract_upload_token_from_response):
                # токен = ЗНАЧЕНИЕ "token" у первого фото, а НЕ ключ словаря!
                first = next(iter(photos.values()))
                if isinstance(first, dict) and first.get("token"):
                    return str(first["token"])
                return str(first)
            if isinstance(j, dict):
                for k in ("token", "file_token", "id", "result"):
                    if k in j:
                        return str(j[k])
        except Exception:
            pass
    return s.strip().strip('"')


def _upload_direct_sync(upload_url: str, path: str, filename: str) -> str:
    with open(path, "rb") as f:
        data = f.read()
    try:
        req = urllib.request.Request(upload_url, data=data,
                                     headers={"Content-Type": "application/octet-stream"},
                                     method="POST")
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.read().decode('utf-8')
    except Exception as e:
        logger.warning(f"Прямая загрузка телом не сработала: {e}")

    boundary = "----maxbot" + os.urandom(8).hex()
    body = (f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="data"; filename="{filename}"\r\n'
            f"Content-Type: application/octet-stream\r\n\r\n").encode('utf-8') \
           + data + f"\r\n--{boundary}--\r\n".encode('utf-8')
    req = urllib.request.Request(upload_url, data=body,
                                 headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
                                 method="POST")
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode('utf-8')


def _download_image_sync(url: str) -> str:
    """Скачивает изображение по ссылке во временный файл.
    Проверяет, что сервер вернул именно картинку, а не HTML-страницу
    (иначе MAX отвергнет её с 'Failed to upload image')."""
    req = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0 (compatible; MAXBot)"})
    with urllib.request.urlopen(req, timeout=90) as r:
        ctype = (r.headers.get("Content-Type") or "").lower()
        data = r.read()
    if len(data) < 1024:
        raise RuntimeError(
            f"Сервер отдал слишком маленький файл ({len(data)} байт), "
            f"Content-Type={ctype!r} — вероятно, страница-заглушка")
    sig = data[:12]
    ext = None
    if sig[:8] == b"\x89PNG\r\n\x1a\n":
        ext = ".png"
    elif sig[:3] == b"\xff\xd8\xff":
        ext = ".jpg"
    elif sig[:6] in (b"GIF87a", b"GIF89a"):
        ext = ".gif"
    elif sig[:4] == b"RIFF" and sig[8:12] == b"WEBP":
        ext = ".webp"
    if not ext:
        head = data[:200].decode("utf-8", errors="replace")
        raise RuntimeError(
            f"По ссылке вернулся не файл изображения, а страница "
            f"(Content-Type={ctype!r}): {head[:120]}")
    fd, path = tempfile.mkstemp(prefix="maxbot_img_", suffix=ext)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    logger.info(f"Скачано изображение: {len(data)} байт, {ctype or ext}")
    return path


async def send_image(chat_id: int, image_path_or_url: str, caption: str = ""):
    if not image_path_or_url:
        logger.warning("Путь/URL изображения не указан")
        return

    try:
        # Ветка URL: сначала пробуем отдать ссылку серверу MAX.
        # Если MAX не смог сам скачать URL — качаем сами и загружаем как файл.
        if image_path_or_url.startswith("http"):
            try:
                att = Attachment(type="image", payload={"url": image_path_or_url})
                await bot.send_message(chat_id=chat_id, text=caption, attachments=[att])
                logger.info(f"Изображение отправлено (URL): {image_path_or_url}")
                return
            except Exception as e_url:
                logger.warning(
                    f"Отправка по URL не удалась ({str(e_url)[:160]}); "
                    f"скачиваю картинку сам и загружаю в MAX как файл")
                tmp_path = await asyncio.to_thread(
                    _download_image_sync, image_path_or_url)
                try:
                    await send_image(chat_id, tmp_path, caption)
                finally:
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
                return

        if not os.path.exists(image_path_or_url):
            logger.warning(f"Файл не найден: {image_path_or_url}")
            return

        filename = os.path.basename(image_path_or_url)
        utype = _pick_upload_type()

        up = await bot.get_upload_url(type=utype)
        up_url = None
        if isinstance(up, str) and up.startswith("http"):
            up_url = up
        else:
            for attr in ("url", "value", "link", "upload_url"):
                v = getattr(up, attr, None)
                if isinstance(v, str) and v.startswith("http"):
                    up_url = v
                    break
        if not up_url:
            raise RuntimeError(f"Не нашли upload URL в ответе: {up!r}")

        with open(image_path_or_url, "rb") as f:
            data = f.read()

        raw_upload = None
        try:
            raw_upload = await bot.upload_file_buffer(filename, up_url, data, utype)
            logger.info(f"Загрузка OK (upload_file_buffer), сырой ответ: {str(raw_upload)[:120]}")
        except Exception as e1:
            logger.warning(f"upload_file_buffer не сработал: {e1}; пробую upload_file")

        if raw_upload is None:
            try:
                raw_upload = await bot.upload_file(up_url, image_path_or_url, utype)
                logger.info(f"Загрузка OK (upload_file), сырой ответ: {str(raw_upload)[:120]}")
            except Exception as e2:
                logger.warning(f"upload_file не сработал: {e2}; прямая загрузка")

        if raw_upload is None:
            raw_upload = await asyncio.to_thread(_upload_direct_sync, up_url,
                                                 image_path_or_url, filename)
            logger.info(f"Загрузка OK (прямая), сырой ответ: {str(raw_upload)[:120]}")

        token = _extract_token(raw_upload)
        if not token:
            raise RuntimeError(f"Токен не извлечён из: {str(raw_upload)[:200]}")

        def _send(tok):
            return bot.send_message(
                chat_id=chat_id, text=caption,
                attachments=[Attachment(type="image", payload={"token": tok})]
            )

        candidates = [token]
        raw_s = str(raw_upload).strip() if raw_upload is not None else ""
        if raw_s.startswith("{") and raw_s != token:
            candidates.append(raw_s)

        last_err: Optional[Exception] = None
        i = 0
        while i < len(candidates):
            cand = candidates[i]
            logger.info(f"Отправка вложения, кандидат токена: {str(cand)[:32]}…")
            for attempt in range(1, 4):
                try:
                    await _send(cand)
                    logger.info(f"Изображение отправлено (файл): {image_path_or_url}")
                    return
                except Exception as e:
                    last_err = e
                    es = str(e)

                    m = re.search(r'photos"\s*:\s*\{\s*"((?:[^"\\]|\\.)*)"', es)
                    if m:
                        key = m.group(1)
                        try:
                            key = key.encode('utf-8').decode('unicode_escape')
                        except Exception:
                            key = key.replace('\\u003d', '=')
                        if key != cand and key not in candidates:
                            logger.info("Сервер подсказал ключ photos — добавляю кандидатом")
                            candidates.append(key)

                    if ('internal.error' in es or 'code=500' in es
                            or "'code': 500" in es or 'Error ID' in es):
                        logger.warning(f"500 (попытка {attempt}/3), пауза 3 с…")
                        await asyncio.sleep(3)
                        continue

                    break
            i += 1

        raise RuntimeError(f"Не удалось отправить изображение; последняя ошибка: {last_err}")

    except Exception as e:
        logger.error(f"Ошибка отправки изображения {image_path_or_url}: {e}")


# ---------- ОСНОВНОЙ ОБРАБОТЧИК ТЕКСТА ----------
@dp.message_created(F.message.body.text)
async def handle_text(event: MessageCreated):
    text = event.message.body.text
    if not text:
        return

    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    username = getattr(user, 'username', None)
    name = getattr(user, 'first_name', 'Пользователь')
    ensure_user(user_id, chat_id, username, name)

    if text.startswith('/'):
        first_word = text.split(maxsplit=1)[0].split('@')[0].lower()
        if first_word not in KNOWN_COMMANDS:
            await event.message.answer(
                "❓ Неизвестная команда. Список доступных команд: /help",
                attachments=[create_main_menu(user_id)]
            )
        return

    state = user_states.get(user_id)

    if state == UserState.WAITING_FOR_FORWARD:
        await process_userinfo_input(event, user_id, chat_id, text)
        return

    # ВЛАДЕЛЕЦ: ввод ID для кнопок ➕/➖ Админы
    if state == UserState.WAITING_ADDADMIN_ID:
        raw = text.strip().strip('`')
        if not raw.isdigit():
            await event.message.answer(
                "⚠️ Нужно число — user_id пользователя.\n"
                "Пример: `65700242`\nДля отмены — /cancel.",
                format="markdown",
                attachments=[create_cancel_menu()]
            )
            return

        new_id = int(raw)
        if new_id == ADMIN_ID:
            await event.message.answer("ℹ️ Владелец уже является администратором.",
                                       attachments=[create_main_menu(user_id)])
            user_states.pop(user_id, None)
            return
        if new_id in _dynamic_admins:
            await event.message.answer(f"ℹ️ Пользователь {new_id} уже администратор.",
                                       attachments=[create_main_menu(user_id)])
            user_states.pop(user_id, None)
            return
        if str(new_id) not in load_users():
            await event.message.answer(
                f"⚠️ Пользователь {new_id} не найден в базе.\n"
                "Он должен хотя бы раз написать боту (/start).",
                attachments=[create_cancel_menu()]
            )
            return

        _dynamic_admins.add(new_id)
        _save_admins()
        user_states.pop(user_id, None)
        logger.info(f"Владелец назначил админа: {new_id}")
        await event.message.answer(
            f"✅ Пользователь `{new_id}` назначен администратором.",
            format="markdown",
            attachments=[create_main_menu(user_id)])
        try:
            await bot.send_message(chat_id=new_id, text="🎉 Вам выданы права администратора бота.")
        except Exception:
            pass
        return

    if state == UserState.WAITING_DELADMIN_ID:
        raw = text.strip().strip('`')
        if not raw.isdigit():
            await event.message.answer(
                "⚠️ Нужно число — user_id администратора.\nДля отмены — /cancel.",
                format="markdown",
                attachments=[create_cancel_menu()]
            )
            return

        del_id = int(raw)
        if del_id == ADMIN_ID:
            await event.message.answer("⛔ Нельзя снять самого владельца.",
                                       attachments=[create_main_menu(user_id)])
            user_states.pop(user_id, None)
            return
        if del_id not in _dynamic_admins:
            await event.message.answer(
                f"ℹ️ Пользователь {del_id} не является динамическим админом.",
                attachments=[create_main_menu(user_id)])
            user_states.pop(user_id, None)
            return

        _dynamic_admins.discard(del_id)
        _save_admins()
        user_states.pop(user_id, None)
        logger.info(f"Владелец снял админа: {del_id}")
        await event.message.answer(
            f"✅ Права администратора сняты с `{del_id}`.",
            format="markdown",
            attachments=[create_main_menu(user_id)])
        try:
            await bot.send_message(chat_id=del_id, text="ℹ️ Права администратора бота отозваны.")
        except Exception:
            pass
        return

    if state == UserState.ADMIN_DEBT_MODE:
        full_name = text.strip()
        if len(full_name.split()) < 2:
            await event.message.answer(
                "⚠️ **Нужно ввести хотя бы фамилию и имя.**\n"
                "Например: *Иванов Иван* или *Иванов Иван Иванович*\n"
                "Вы можете нажать «Отмена» для выхода.",
                format="markdown",
                attachments=[create_cancel_menu()]
            )
            return

        user_states.pop(user_id, None)
        await process_admin_debt_search(user_id, chat_id, full_name)
        return

    if state == UserState.ADMIN_DEBT_CLARIFY_MODE:
        data = admin_debt_data.get(user_id)
        if not data:
            user_states.pop(user_id, None)
            await event.message.answer(
                "❌ **Ошибка: данные для уточнения не найдены.**\n"
                "Начните проверку заново.",
                attachments=[create_main_menu(user_id)]
            )
            return

        lw = [_norm_word(w) for w in text.strip().split()]
        base_words = str(data['base']).split()
        variants: Dict[str, str] = data['variants']

        candidates: List[str] = []
        if len(lw) == 1:
            candidates.append(' '.join(base_words + lw))
        else:
            candidates.append(' '.join(lw[:3]))
            candidates.append(' '.join(lw[:2]))

        found_key = next((c for c in candidates if c in variants), None)
        if found_key is None:
            entered_norm = ' '.join(lw[:3])
            for key, display in variants.items():
                if ' '.join(_norm_word(w) for w in display.split()[:3]) == entered_norm:
                    found_key = key
                    break

        if found_key:
            matches = await search_debt_with_recovery(found_key)
            user_states.pop(user_id, None)
            admin_debt_data.pop(user_id, None)
            if matches:
                groups = group_matches(matches)
                matches = groups.get(found_key, matches)
                await send_debt_result(chat_id, user_id, matches, menu=create_admin_debt_menu())
            else:
                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ **Не удалось найти долг по уточнённому ФИО.**",
                    attachments=[create_admin_debt_menu()]
                )
        else:
            await event.message.answer(
                "❌ **Вариант не найден.**\n\n"
                "Проверьте ввод и попробуйте снова.\n"
                "Доступные варианты:\n" + "\n".join([f"• {n}" for n in variants.values()]),
                format="markdown",
                attachments=[create_cancel_menu()]
            )
        return

    if state == UserState.DEBT_CLARIFY_MODE:
        entered = text.strip()
        data = user_clarify_data.get(user_id)
        if not data:
            user_states.pop(user_id, None)
            await event.message.answer(
                "❌ **Ошибка: данные для уточнения не найдены.**\n"
                "Пожалуйста, начните проверку долга заново.",
                attachments=[create_main_menu(user_id)]
            )
            return

        lw = [_norm_word(w) for w in entered.split()]
        base_words = str(data['base']).split()
        variants: Dict[str, str] = data['variants']

        candidates: List[str] = []
        if len(lw) == 1:
            candidates.append(' '.join(base_words + lw))
        else:
            candidates.append(' '.join(lw[:3]))
            candidates.append(' '.join(lw[:2]))

        found_key = next((c for c in candidates if c in variants), None)
        if found_key is None:
            entered_norm = ' '.join(lw[:3])
            for key, display in variants.items():
                if ' '.join(_norm_word(w) for w in display.split()[:3]) == entered_norm:
                    found_key = key
                    break

        if found_key:
            display_name = variants[found_key]
            set_user_full_name(user_id, display_name)
            matches = await search_debt_with_recovery(found_key)
            user_states.pop(user_id, None)
            user_clarify_data.pop(user_id, None)
            if matches:
                groups = group_matches(matches)
                matches = groups.get(found_key, matches)
                await send_debt_result(chat_id, user_id, matches)
            else:
                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ **Не удалось найти долг по уточнённому ФИО.**\n"
                         "Пожалуйста, обратитесь к администратору.",
                    attachments=[create_main_menu(user_id)]
                )
        else:
            await event.message.answer(
                "❌ **Вариант не найден.**\n\n"
                "Пожалуйста, проверьте правильность ввода и попробуйте снова.\n"
                "Доступные варианты:\n" + "\n".join([f"• {n}" for n in variants.values()]),
                format="markdown",
                attachments=[create_cancel_menu()]
            )
        return

    if state == UserState.DEBT_NAME_MODE:
        full_name = text.strip()
        if len(full_name.split()) < 2:
            await event.message.answer(
                "⚠️ **Нужно ввести хотя бы фамилию и имя.**\n"
                "Например: *Иванов Иван* или *Иванов Иван Иванович*\n"
                "Вы можете нажать «Отмена» для выхода.",
                format="markdown",
                attachments=[create_cancel_menu()]
            )
            return

        set_user_full_name(user_id, full_name)
        user_states.pop(user_id, None)

        await process_debt_search(user_id, chat_id, full_name)
        return

    if state == UserState.BROADCAST_MODE:
        if not is_admin(user_id):
            user_states.pop(user_id, None)
            await event.message.answer(
                "⛔ Доступно только администратору.",
                attachments=[create_main_menu(user_id)]
            )
            return

        users = load_users()
        sent_count = 0
        error_count = 0
        broadcast_ok: List[tuple] = []
        logger.info(f"Начинаем рассылку от админа {user_id}, пользователей: {len(users)}")

        for uid, user_data in users.items():
            chat_id_user = user_data.get('chat_id')
            if not chat_id_user:
                continue
            try:
                await bot.send_message(
                    chat_id=chat_id_user,
                    text=f"📢 **Новость от администратора:**\n\n{text}",
                    format="markdown"
                )
                sent_count += 1
                broadcast_ok.append((uid, user_data))
            except Exception as e:
                logger.error(f"Не удалось отправить новость пользователю {uid}: {e}")
                error_count += 1
            await asyncio.sleep(0.2)

        _log_sent_bulk(broadcast_ok, "📢 Рассылка новостей",
                       f"📢 **Новость от администратора:**\n\n{text}")
        user_states.pop(user_id, None)
        await event.message.answer(
            f"✅ **Рассылка завершена.**\n"
            f"Отправлено: {sent_count}\n"
            f"Ошибок: {error_count}",
            format="markdown",
            attachments=[create_main_menu(user_id)]
        )
        return

    if is_admin(user_id):
        await event.message.answer(
            "👋 Вы администратор. Используйте:\n"
            "• /menu или /start – показать меню с кнопками\n"
            "• /broadcast – для рассылки новостей\n"
            "• /sentlog – журнал отправок бота (админ)\n"
            "• /userinfo – для получения данных пользователя\n"
            "• /debt или кнопка «Долг» – проверить задолженность по ФИО клиента\n"
            "• /refresh_debt – проверить таблицы сверки сейчас\n"
            "• /debt_status – статус сверки",
            format="markdown",
            attachments=[create_main_menu(user_id)]
        )
        return

    await event.message.answer(
        "Не понял вас. Используйте кнопки под сообщениями или /help для справки.",
        attachments=[create_main_menu(user_id)]
    )


# ---------- ОБРАБОТЧИК СООБЩЕНИЙ БЕЗ ТЕКСТА ----------
@dp.message_created()
async def handle_non_text(event: MessageCreated):
    user = event.message.sender
    user_id = user.user_id
    chat_id = event.message.recipient.chat_id
    username = getattr(user, 'username', None)
    name = getattr(user, 'first_name', 'Пользователь')

    body = getattr(event.message, 'body', None)
    text = getattr(body, 'text', '') if body else ''

    if text:
        return

    ensure_user(user_id, chat_id, username, name)
    state = user_states.get(user_id)

    if state == UserState.WAITING_FOR_FORWARD:
        await process_userinfo_input(event, user_id, chat_id, text or "")
        return

    if state in (UserState.ADMIN_DEBT_MODE, UserState.ADMIN_DEBT_CLARIFY_MODE):
        await event.message.answer(
            "⚠️ Введите ФИО текстом.",
            attachments=[create_cancel_menu()]
        )
        return

    logger.info(f"Сообщение без текста от user_id={user_id}, state={state} — игнор")


# ==================== ЗАПУСК БОТА ====================

async def main():
    logger.info("Запуск бота...")

    # Регистрируем команды на сервере MAX; результат логируем для диагностики
    res = await bot.set_commands(
        BotCommand(name="start", description="Главное меню"),
        BotCommand(name="menu", description="Главное меню"),
        BotCommand(name="help", description="Справка"),
        BotCommand(name="debt", description="Долг"),
        BotCommand(name="change_name", description="Сменить ФИО"),
        BotCommand(name="userinfo", description="ID контакта (админ)"),
        BotCommand(name="broadcast", description="Рассылка (только админ)"),
        BotCommand(name="sentlog", description="Журнал отправок (админ)"),
        BotCommand(name="refresh_debt", description="Обновить сверку (админ)"),
        BotCommand(name="debt_status", description="Статус сверки (админ)"),
        BotCommand(name="adminlist", description="Список админов"),
        BotCommand(name="cancel", description="Отменить текущее действие"),
    )
    logger.info(f"set_commands результат: {res!r}")

    await bot.delete_webhook()

    # Фоновый цикл: стартовая загрузка + расписание пн/вт-пт/сб-вс
    asyncio.create_task(debt_refresh_loop())

    # Не обрабатывать накопившиеся за офлайн-время события
    await dp.start_polling(bot, skip_updates=True)


if __name__ == "__main__":
    asyncio.run(main())
