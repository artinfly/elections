import re
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Iterator, Optional

import openpyxl
from django.db.models import Q
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .models import DEG, METHOD_LABELS, UIK, UIK19, UVZ

# Константы
BAD_FORMAT = "Ошибка, документ не соответствует формату"
BATCH = 500
NO_PRODUCTION = "Без производства"

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


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _date(value: Any) -> Optional[datetime.date]:
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
    name = (value or "").strip()
    return (0, int(name), "") if name.isdigit() else (1, 0, name)


def padded_number(value: Any) -> str:
    name = (value or "").strip()
    return f"{int(name):03d}" if name.isdigit() else name


def department_file_name(department: str) -> str:
    return BAD_NAME_CHARS.sub("-", padded_number(department))


def _apply_border(sheet: Any) -> None:
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row_cells in sheet.iter_rows(
        min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column
    ):
        for cell in row_cells:
            cell.border = border


def _format_with_percent(count: int, total: int) -> str:
    if not total:
        return str(count)
    return f"{count} ({round(count / total * 100)}%)"
