from datetime import datetime

import openpyxl
from django.utils import timezone
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from utils.archiver import ReportArchiver

from .models import METHOD_LABELS, Employee

BAD_FORMAT = "Ошибка, документ не соответствует формату"

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


def _text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _date(value):
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


def _open(upload):
    try:
        book = openpyxl.load_workbook(upload, read_only=True, data_only=True)
    except Exception:
        raise ValueError(BAD_FORMAT)
    return book.active.iter_rows(values_only=True)


def _header(rows, wanted):
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


BATCH = 500


def _rows_by_tab(rows, positions):
    parsed = {}
    for row in rows:
        if not any(row):
            continue
        values = {}
        for name, index in positions.items():
            field = COLUMNS[name]
            cell = row[index] if index < len(row) else None
            values[field] = _date(cell) if field == "birth_date" else _text(cell)
        tab = values.pop("tab_number")
        if tab:
            parsed[tab] = values
    return parsed


def _known_rows(tabs, fields):
    known = {}
    tabs = list(tabs)
    for start in range(0, len(tabs), 2000):
        chunk = tabs[start : start + 2000]
        for row in Employee.objects.filter(tab_number__in=chunk).values(
            "pk", "tab_number", *fields
        ):
            known[row.pop("tab_number")] = row
    return known


def import_base(upload):
    rows = _open(upload)
    positions = _header(rows, COLUMNS)
    if "Таб№" not in positions:
        raise ValueError(BAD_FORMAT)

    parsed = _rows_by_tab(rows, positions)
    if not parsed:
        return 0, 0, 0

    fields = [COLUMNS[name] for name in positions if COLUMNS[name] != "tab_number"]
    known = _known_rows(parsed, fields)

    fresh, stale = [], []
    for tab, values in parsed.items():
        current = known.get(tab)
        if current is None:
            fresh.append(Employee(tab_number=tab, **values))
        elif any(current[field] != values[field] for field in fields):
            stale.append(Employee(pk=current["pk"], tab_number=tab, **values))

    if fresh:
        Employee.objects.bulk_create(fresh, batch_size=BATCH)
    if stale:
        Employee.objects.bulk_update(stale, fields, batch_size=BATCH)
    return len(fresh), len(stale), len(parsed)


def set_turnout(queryset, voted=True):
    fields = {"voted": voted, "voted_at": timezone.now() if voted else None}
    if not voted:
        fields["voted_method"] = ""
    return queryset.update(**fields)


def mark_voted(tabs, voted=True):
    tabs = {t for t in tabs if t}
    if not tabs:
        return 0, 0
    found = Employee.objects.filter(tab_number__in=tabs)
    missing = len(tabs) - found.count()
    return set_turnout(found, voted), missing


REPORT_WIDTHS = (14, 46, 10)


def department_report(department, moment=None):
    moment = moment or timezone.localtime()
    people = Employee.objects.filter(department=department)
    total = people.count()
    voted = people.filter(voted=True).count()

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = f"Цех {department}"[:31]
    for index, width in enumerate(REPORT_WIDTHS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    bold = Font(bold=True)
    title = sheet.cell(1, 1, f"Информация по голосованию на {moment:%d.%m.%y %H:%M}")
    title.font = Font(bold=True, size=12)
    sheet.cell(2, 1, f"Цех {department}").font = bold

    for column, name in enumerate(
        ("Общее количество голосующих", "Количество проголосовавших", "Процент"), 1
    ):
        cell = sheet.cell(4, column, name)
        cell.font = bold
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    sheet.cell(5, 1, total).alignment = Alignment(horizontal="center")
    sheet.cell(5, 2, voted).alignment = Alignment(horizontal="center")
    share = sheet.cell(5, 3, (voted / total) if total else 0)
    share.number_format = "0.00%"
    share.alignment = Alignment(horizontal="center")

    sheet.cell(7, 1, "Список НЕ принявших участие в голосовании:").font = bold

    row = 8
    for person in people.filter(voted=False).order_by("surname", "name", "patronymic"):
        sheet.cell(row, 1, person.tab_number)
        sheet.cell(row, 2, person.fio)
        sheet.cell(row, 3, 1).alignment = Alignment(horizontal="center")
        row += 1

    return book


def reports_archive(moment=None):
    moment = moment or timezone.localtime()
    departments = (
        Employee.objects.exclude(department="")
        .values_list("department", flat=True)
        .distinct()
        .order_by("department")
    )
    archiver = ReportArchiver()
    for department in departments:
        archiver.add_workbook(
            department_report(department, moment), f"{department}.xlsx"
        )
    return archiver


def export_xlsx():
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сотрудники"
    sheet.append(list(COLUMNS) + ["Способ (план)", "Проголосовал", "Где голосовал"])
    fields = [COLUMNS[name] for name in COLUMNS]
    for person in Employee.objects.all().iterator(chunk_size=2000):
        row = [getattr(person, f) for f in fields]
        row.append(METHOD_LABELS.get(person.method, ""))
        row.append("да" if person.voted else "нет")
        row.append(METHOD_LABELS.get(person.voted_method, ""))
        sheet.append(row)
    return book
