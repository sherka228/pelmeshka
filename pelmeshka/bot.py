import os
import re
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

# ============================================================
# НАСТРОЙКИ
# ============================================================

BOT_TOKEN = "8800460586:AAE4ZVC_YSGflWqJOS-RUFwDUN1LfLRNve8"
# Excel-файл должен лежать рядом с bot.py
EXCEL_FILE = "расписание_на_5_неделю_с_28_09_по_02_10_2026_г.xlsx"

# Если нужно показывать расписание конкретной группы,
# можно указать номер группы здесь, например "5301".
# Если оставить None, бот будет использовать выбранный профиль.
DEFAULT_GROUP = None


# ============================================================
# СОСТОЯНИЯ
# ============================================================

class Setup(StatesGroup):
    choosing_course = State()
    choosing_profile = State()


# ============================================================
# РАБОТА С EXCEL
# ============================================================

DAY_NAMES = {
    "ПОНЕДЕЛЬНИК": "Понедельник",
    "ВТОРНИК": "Вторник",
    "СРЕДА": "Среда",
    "ЧЕТВЕРГ": "Четверг",
    "ПЯТНИЦА": "Пятница",
    "СУББОТА": "Суббота",
    "ВОСКРЕСЕНЬЕ": "Воскресенье",
}
DAYS_BY_NUM = {
    0: "Понедельник",
    1: "Вторник",
    2: "Среда",
    3: "Четверг",
    4: "Пятница",
    5: "Суббота",
    6: "Воскресенье",
}


def clean_text(value):
    """Очищает текст из Excel."""
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*", "\n", text)
    return text.strip()


def parse_date(value):
    """Преобразует дату из Excel в date."""
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if value is None:
        return None

    text = clean_text(value)

    # Например: 29.09.2026 г.
    match = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    if match:
        day, month, year = map(int, match.groups())
        return date(year, month, day)

    return None


def find_date_column(ws):
    """
    В файле первых курсов дата находится в C,
    у 3 курса — в B. Поэтому определяем колонку автоматически.
    """
    for row in range(1, min(ws.max_row, 10) + 1):
        for col in range(1, min(ws.max_column, 8) + 1):
            if clean_text(ws.cell(row, col).value).lower() == "дата":
                return col

    # Запасной вариант
    return 3


def find_header_row(ws):
    """Ищет строку с заголовками профилей."""
    for row in range(1, min(ws.max_row, 10) + 1):
        values = [
            clean_text(ws.cell(row, col).value).lower()
            for col in range(1, ws.max_column + 1)
        ]

        if any("профиль" in value for value in values):
            return row

    return 4


def get_profiles(ws):
    """
    Возвращает профили и связанные с ними колонки.
    В Excel структура обычно такая:

    профиль | ауд | профиль | ауд | профиль | ауд ...
    """
    header_row = find_header_row(ws)
    profiles = []

    for col in range(1, ws.max_column + 1):
        value = clean_text(ws.cell(header_row, col).value)

        if "профиль" not in value.lower():
            continue

        # Следующая колонка обычно является аудиторией
        room_col = col + 1 if col + 1 <= ws.max_column else None

        # Строка ниже содержит номер группы
        group_value = ""
        if header_row + 1 <= ws.max_row:
            group_value = clean_text(ws.cell(header_row + 1, col).value)

        profiles.append({
            "name": value.replace("профиль", "").strip(" :\n"),
            "column": col,
            "room_column": room_col,
            "group": group_value,
        })

    return profiles


def load_schedule():
    """Загружает Excel и строит удобную структуру расписания."""
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(
            f"Не найден Excel-файл: {EXCEL_FILE}"
        )

    wb = load_workbook(EXCEL_FILE, data_only=True)

    result = {}

    for sheet_name in wb.sheetnames:
        # Лист с заочным обучением пока не используем,
        # потому что формат там может отличаться.
        if "ОЗФ" in sheet_name:
            continue

        ws = wb[sheet_name]
        date_col = find_date_column(ws)
        header_row = find_header_row(ws)
        profiles = get_profiles(ws)

        if not profiles:
            continue

        course = re.sub(r"\s+", " ", sheet_name).strip()

        result[course] = {
            "sheet": ws,
            "date_col": date_col,
            "header_row": header_row,
            "profiles": profiles,
        }

    return result


SCHEDULE = load_schedule()


def get_course_names():
    return list(SCHEDULE.keys())


def get_profile_name(profile):
    name = clean_text(profile["name"])
    group = clean_text(profile["group"])

    # Убираем количество студентов, например:
    # (14 ст.), (25 ст.), (10 ст.)
    name = re.sub(r"\s*\(\d+\s*ст\.\)", "", name)
    group = re.sub(r"\s*\(\d+\s*ст\.\)", "", group)

    if group:
        return f"{name} — {group}"

    return name

def get_schedule_for_date(course, profile_index, target_date):
    """
    Возвращает пары выбранного профиля на конкретную дату.
    """
    course_data = SCHEDULE[course]
    ws = course_data["sheet"]
    date_col = course_data["date_col"]
    profile = course_data["profiles"][profile_index]

    subject_col = profile["column"]
    room_col = profile["room_column"]

    current_date = None
    current_day = ""

    lessons = []

    for row in range(course_data["header_row"] + 2, ws.max_row + 1):
        date_value = ws.cell(row, date_col).value
        parsed = parse_date(date_value)

        if parsed:
            current_date = parsed
            current_day = clean_text(ws.cell(row, date_col + 1).value)

        if current_date != target_date:
            continue

        lesson = clean_text(ws.cell(row, subject_col).value)
        room = clean_text(ws.cell(row, room_col).value) if room_col else ""

        # Время находится рядом с колонкой "Пара".
        # На 1/2/4 курсах обычно date_col + 3,
        # у 3 курса date_col + 3 тоже соответствует времени.
        time_value = clean_text(ws.cell(row, date_col + 3).value)
        pair_value = clean_text(ws.cell(row, date_col + 2).value)

        if not lesson:
            continue

        # Не показываем полностью пустые строки.
        lessons.append({
            "pair": pair_value,
            "time": time_value,
            "subject": lesson,
            "room": room,
        })

    return current_day, lessons


def find_notes(course, target_date):
    """
    Ищет заметки/примечания в строках Excel.
    Если в файле появятся слова 'Примечание', 'Заметка' и т.п.,
    бот сможет их вывести.
    """
    course_data = SCHEDULE[course]
    ws = course_data["sheet"]

    notes = []
    current_date = None

    note_words = (
        "примечание",
        "примечания",
        "заметка",
        "заметки",
        "важно",
        "внимание",
    )

    for row in range(1, ws.max_row + 1):
        # Проверяем всю строку
        values = [
            clean_text(ws.cell(row, col).value)
            for col in range(1, ws.max_column + 1)
        ]

        for value in values:
            parsed = parse_date(value)
            if parsed:
                current_date = parsed
                break

        if current_date != target_date:
            continue

        for value in values:
            lower = value.lower()
            if value and any(word in lower for word in note_words):
                notes.append(value)

    # Убираем дубликаты
    unique_notes = []
    for note in notes:
        if note not in unique_notes:
            unique_notes.append(note)

    return unique_notes


# ============================================================
# КЛАВИАТУРЫ TELEGRAM
# ============================================================

def main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="📅 Сегодня"),
                KeyboardButton(text="📅 Завтра"),
            ],
            [
                KeyboardButton(text="⚙️ Настройки"),
            ],
        ],
        resize_keyboard=True,
    )


def course_keyboard():
    buttons = []

    for course in get_course_names():
        buttons.append([KeyboardButton(text=course)])

    return ReplyKeyboardMarkup(
        keyboard=buttons,
        resize_keyboard=True,
        one_time_keyboard=True,
    )


def profile_keyboard(course):
    profiles = SCHEDULE[course]["profiles"]

    buttons = []

    for i, profile in enumerate(profiles):
        name = get_profile_name(profile)

        # Telegram ограничивает удобство слишком длинных кнопок,
        # поэтому оставляем первые ~55 символов.
        if len(name) > 55:
            name = name[:52] + "..."

        buttons.append([
            KeyboardButton(text=f"{i}: {name}")
        ])

    return ReplyKeyboardMarkup(
        keyboard=buttons,
        resize_keyboard=True,
        one_time_keyboard=True,
    )


# ============================================================
# ФОРМИРОВАНИЕ СООБЩЕНИЯ
# ============================================================

def build_schedule_message(
    course,
    profile_index,
    target_date,
    title=None,
):
    profile = SCHEDULE[course]["profiles"][profile_index]
# Игнорируем имя дня из Excel (первая переменная _), так как определяем его по target_date
    _, lessons = get_schedule_for_date(
        course,
        profile_index,
        target_date,
    )

    notes = find_notes(course, target_date)

    if title is None:
        title = "📅 Расписание на сегодня"

    profile_name = get_profile_name(profile)

    # Точно вычисляем день недели по целевой дате (0 — Понедельник, 6 — Воскресенье)
    pretty_day = DAYS_BY_NUM[target_date.weekday()]

    lines = [
        f"<b>{title}</b>",
        f"📆 {target_date.strftime('%d.%m.%Y')}",
        f"🗓 {pretty_day}",
        f"🎓 {course}",
        f"👥 {profile_name}",
        "",
    ]

    if not lessons:
        lines.append("🎉 На этот день занятий нет.")
    else:
        for lesson in lessons:
            pair = lesson["pair"] or "—"
            time = lesson["time"] or "Время не указано"
            subject = lesson["subject"]
            room = lesson["room"]

            lines.append(f"<b>{pair} пара</b> · {time}")
            lines.append(f"📚 {subject}")

            if room:
                lines.append(f"📍 Аудитория: {room}")

            lines.append("")

    if notes:
        lines.append("📝 <b>Заметки:</b>")
        for note in notes:
            lines.append(f"• {note}")

    return "\n".join(lines)
   

# ============================================================
# TELEGRAM BOT
# ============================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


@dp.message(CommandStart())
async def start(message: Message, state: FSMContext):
    await state.clear()

    await message.answer(
        "Привет! 👋\n\n"
        "Это бот для просмотра учебного расписания.\n"
        "Сначала выбери свой курс:",
        reply_markup=course_keyboard(),
    )

    await state.set_state(Setup.choosing_course)


@dp.message(Command("help"))
async def help_command(message: Message):
    await message.answer(
        "Доступные действия:\n\n"
        "📅 Сегодня — расписание на текущую дату\n"
        "📅 Завтра — расписание на следующий день\n"
        "⚙️ Настройки — изменить курс или профиль\n\n"
        "Чтобы заново настроить бота, используй /start.",
        reply_markup=main_keyboard(),
    )


@dp.message(Setup.choosing_course)
async def choose_course(message: Message, state: FSMContext):
    course = message.text.strip()

    if course not in SCHEDULE:
        await message.answer(
            "Пожалуйста, выбери курс кнопкой ниже.",
            reply_markup=course_keyboard(),
        )
        return

    await state.update_data(course=course)

    await message.answer(
        "Теперь выбери свой профиль / группу:",
        reply_markup=profile_keyboard(course),
    )

    await state.set_state(Setup.choosing_profile)


@dp.message(Setup.choosing_profile)
async def choose_profile(message: Message, state: FSMContext):
    data = await state.get_data()
    course = data.get("course")

    if course not in SCHEDULE:
        await message.answer(
            "Настройка сбилась. Нажми /start.",
            reply_markup=main_keyboard(),
        )
        await state.clear()
        return

    match = re.match(r"^(\d+):", message.text.strip())

    if not match:
        await message.answer(
            "Пожалуйста, выбери профиль кнопкой.",
            reply_markup=profile_keyboard(course),
        )
        return

    profile_index = int(match.group(1))

    if profile_index >= len(SCHEDULE[course]["profiles"]):
        await message.answer(
            "Такого профиля нет. Выбери профиль ещё раз.",
            reply_markup=profile_keyboard(course),
        )
        return

    USER_SETTINGS[message.from_user.id] = {
        "course": course,
        "profile_index": profile_index,
    }

    await state.clear()

    await message.answer(
        "Готово! ✅\n\n"
        "Теперь нажми «📅 Сегодня», чтобы получить расписание.",
        reply_markup=main_keyboard(),
    )


# Настройки пользователей.
# Для простого бота этого достаточно.
# После перезапуска бота выбор пользователей сбросится.
USER_SETTINGS = {}


async def send_schedule(message: Message, target_date: date, title: str):
    user_id = message.from_user.id
    data = USER_SETTINGS.get(user_id, {})

    course = data.get("course")
    profile_index = data.get("profile_index")

    if course is None or profile_index is None:
        await message.answer(
            "Сначала нужно выбрать курс и профиль. Нажми /start."
        )
        return

    text = build_schedule_message(
        course=course,
        profile_index=profile_index,
        target_date=target_date,
        title=title,
    )

    await message.answer(
        text,
        parse_mode="HTML",
        reply_markup=main_keyboard(),
    )


@dp.message(F.text == "📅 Сегодня")
async def today(message: Message):
    await send_schedule(
        message,
        datetime.now(ZoneInfo("Asia/Tashkent")).date(),
        "📅 Расписание на сегодня",
    )


@dp.message(F.text == "📅 Завтра")
async def tomorrow(message: Message):
    await send_schedule(
        message,
        datetime.now(ZoneInfo("Asia/Tashkent")).date() + timedelta(days=1),
        "📅 Расписание на завтра",
    )


@dp.message(F.text == "⚙️ Настройки")
async def settings(message: Message, state: FSMContext):
    await state.clear()

    await message.answer(
        "Выбери курс:",
        reply_markup=course_keyboard(),
    )

    await state.set_state(Setup.choosing_course)


async def main():
    if BOT_TOKEN == "ВСТАВЬ_СЮДА_ТОКЕН_БОТА":
        raise RuntimeError(
            "Укажи токен бота в переменной BOT_TOKEN."
        )

    print("Бот запущен.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
