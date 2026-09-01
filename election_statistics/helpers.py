"""
Модуль вспомогательных функций и констант.
Содержит утилиты для парсинга Excel, форматирования данных и настройки отчетов,
не зависящие от сложной бизнес-логики.
"""

import re
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Iterator, Optional

import openpyxl
from django.db.models import Q
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .models import DEG, UIK, UIK19, UVZ

# ==============================================================================
# Константы
# ==============================================================================

BAD_FORMAT = "Ошибка: документ не соответствует ожидаемому формату"
BATCH = 500  # Размер пачки для bulk_create / bulk_update
NO_PRODUCTION = "Без производства"

# Соответствие заголовков колонок старого формата Excel полям модели Employee
COLUMNS = {
    "Подразделение": "department",
    "Таб№": "tab_number",
    "Фамилия": "surname",
    "Имя": "name",
    "Отчество": "patronymic",
    "Должность": "position",
    "Категория": "category",
    "Дата рождения": "birth_date",
    "Регион": "region",
    "Город": "city",
    "Улица": "street",
    "Дом": "house",
    "УИК": "uik",
    "Адрес УИК": "uik_address",
    "Район": "district",
    "Округ": "okrug",
}

# Настройки режимов отчета по цеху (явка или выбор способа)
REPORT_MODES = {
    "turnout": {
        "title": "Информация по голосованию на",
        "column": "Количество проголосовавших",
        "list_title": "Список НЕ принявших участие в голосовании:",
        "done": Q(voted=True),
    },
    "method": {
        "title": "Информация по выбору способа голосования на",
        "column": "Количество выбравших способ",
        "list_title": "Список НЕ выбравших способ голосования:",
        "done": ~Q(method=""),
    },
}

BAD_NAME_CHARS = re.compile(r"[\\/*?:\[\]]")


# ==============================================================================
# Утилиты парсинга и форматирования
# ==============================================================================


def _text(value: Any) -> str:
    """Приводит значение ячейки Excel к строке. Float без дробной части превращает в int."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _date(value: Any) -> Optional[datetime.date]:
    """Разбирает дату из ячейки Excel. Поддерживает datetime, date и строки 'дд.мм.гггг' / 'гггг-мм-дд'."""
    if isinstance(value, datetime):
        return value.date()
    if hasattr(value, "year"):
        return value
    head = _text(value).split()
    if not head:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(head[0], fmt).date()
        except ValueError:
            continue
    return None


@contextmanager
def _sheet(upload: Any) -> Iterator:
    """
    Контекст-менеджер для безопасного чтения Excel.
    Открывает в read_only режиме и гарантированно закрывает ресурсы.
    """
    try:
        book = openpyxl.load_workbook(upload, read_only=True, data_only=True)
    except Exception:
        raise ValueError(BAD_FORMAT)

    rows = book.active.iter_rows(values_only=True)
    try:
        yield rows
    finally:
        rows.close()
        book.close()


def _header(rows: Iterator, wanted: dict) -> dict:
    """Ищет строку с заголовками. Возвращает словарь {имя_колонки: индекс}. Регистронезависимый поиск."""
    lookup = {name.casefold(): name for name in wanted}
    for row in rows:
        found = {}
        for index, cell in enumerate(row):
            key = _text(cell).casefold()
            if key in lookup:
                found[lookup[key]] = index
        if found:
            return found
    raise ValueError(BAD_FORMAT)


def _by_number(value: Any) -> tuple:
    """Ключ сортировки: числовые значения идут первыми по возрастанию, строки — после в алфавитном порядке."""
    name = (value or "").strip()
    return (0, int(name), "") if name.isdigit() else (1, 0, name)


def padded_number(value: Any) -> str:
    """Дополняет числовой номер цеха нулями до 3 знаков (например, '7' -> '007')."""
    name = (value or "").strip()
    return f"{int(name):03d}" if name.isdigit() else name


def department_file_name(department: str) -> str:
    """Формирует безопасное имя файла для отчета, удаляя запрещенные символы."""
    return BAD_NAME_CHARS.sub("-", padded_number(department))


def _apply_border(sheet: Any) -> None:
    """Применяет тонкую черную рамку ко всем заполненным ячейкам листа."""
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row_cells in sheet.iter_rows(
        min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column
    ):
        for cell in row_cells:
            cell.border = border


def _format_with_percent(count: int, total: int) -> str:
    """Форматирует число с процентом от общего количества (например, '5 (50%)')."""
    if not total:
        return str(count)
    return f"{count} ({round(count / total * 100)}%)"
