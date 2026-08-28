import re
from contextlib import contextmanager
from datetime import datetime

import openpyxl
from django.db.models import Count, F, Q
from django.utils import timezone
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from utils.archiver import ReportArchiver

from .models import DEG, METHOD_LABELS, UIK, UIK19, UVZ, Employee

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


@contextmanager
def _sheet(upload):
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
    with _sheet(upload) as rows:
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
    return queryset.update(
        voted=voted,
        voted_at=timezone.now() if voted else None,
        voted_method=F("method") if voted else "",
    )


def mark_voted(tabs, voted=True):
    tabs = {t for t in tabs if t}
    if not tabs:
        return 0, 0
    found = Employee.objects.filter(tab_number__in=tabs)
    missing = len(tabs) - found.count()
    return set_turnout(found, voted), missing


REPORT_WIDTHS = (14, 46, 10)

BAD_NAME_CHARS = re.compile(r"[\\/*?:\[\]]")

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


def padded_number(value):
    name = (value or "").strip()
    return f"{int(name):03d}" if name.isdigit() else name


def department_file_name(department):
    return BAD_NAME_CHARS.sub("-", padded_number(department))


def department_report(department, moment=None, mode="turnout"):
    rule = REPORT_MODES[mode]
    moment = moment or timezone.localtime()
    people = Employee.objects.filter(department=department)
    total = people.count()
    done = people.filter(rule["done"]).count()

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = BAD_NAME_CHARS.sub("-", f"Цех {department}")[:31]
    for index, width in enumerate(REPORT_WIDTHS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    bold = Font(bold=True)
    title = sheet.cell(1, 1, f"{rule['title']} {moment:%d.%m.%y %H:%M}")
    title.font = Font(bold=True, size=12)
    sheet.cell(2, 1, f"Цех {department}").font = bold

    for column, name in enumerate(
        ("Общее количество голосующих", rule["column"], "Процент"), 1
    ):
        cell = sheet.cell(4, column, name)
        cell.font = bold
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    sheet.cell(5, 1, total).alignment = Alignment(horizontal="center")
    sheet.cell(5, 2, done).alignment = Alignment(horizontal="center")
    share = sheet.cell(5, 3, (done / total) if total else 0)
    share.number_format = "0.00%"
    share.alignment = Alignment(horizontal="center")

    sheet.cell(7, 1, rule["list_title"]).font = bold

    row = 8
    for person in people.exclude(rule["done"]).order_by(
        "surname", "name", "patronymic"
    ):
        sheet.cell(row, 1, person.tab_number)
        sheet.cell(row, 2, person.fio)
        row += 1

    return book


def reports_archive(moment=None, mode="turnout"):
    moment = moment or timezone.localtime()
    departments = (
        Employee.objects.exclude(department="")
        .values_list("department", flat=True)
        .distinct()
        .order_by()
    )
    archiver = ReportArchiver()
    taken = set()
    for department in sorted(departments, key=_by_number):
        name = department_file_name(department)
        unique, attempt = name, 2
        while unique in taken:
            unique = f"{name}-{attempt}"
            attempt += 1
        taken.add(unique)
        archiver.add_workbook(
            department_report(department, moment, mode), f"{unique}.xlsx"
        )
    return archiver


SUMMARY_WIDTHS = (18, 16, 10, 10, 12, 12, 16, 10)


def _by_number(value):
    name = (value or "").strip()
    return (0, int(name), "") if name.isdigit() else (1, 0, name)


NO_PRODUCTION = "Без производства"
PRODUCTION_WIDTHS = (16, 14, 16, 10)


def _share_row(sheet, line, label, people, came, bold=None):
    cells = (
        sheet.cell(line, 1, label),
        sheet.cell(line, 2, people),
        sheet.cell(line, 3, came),
        sheet.cell(line, 4, came / people if people else 0),
    )
    for cell in cells[1:]:
        cell.alignment = Alignment(horizontal="center")
    cells[3].number_format = "0.00%"
    if bold:
        for cell in cells:
            cell.font = bold
    return line + 1


def production_table():
    grouped = {}
    for row in (
        Employee.objects.exclude(department="")
        .values("production", "department")
        .annotate(people=Count("id"), came=Count("id", filter=Q(voted=True)))
        .order_by()
    ):
        grouped.setdefault(row["production"] or NO_PRODUCTION, []).append(row)

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "По производствам"
    for index, width in enumerate(PRODUCTION_WIDTHS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    bold = Font(bold=True)
    sheet.merge_cells("A1:A2")
    sheet.merge_cells("B1:B2")
    sheet.merge_cells("C1:D1")
    for coordinate, title in (
        ("A1", "Подразделение"),
        ("B1", "Общее число работающих"),
        ("C1", "Итог"),
        ("C2", "Количество проголосовавших"),
        ("D2", "%"),
    ):
        cell = sheet[coordinate]
        cell.value = title
        cell.font = bold
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    line = 3
    total_people = total_came = 0
    for production in sorted(grouped, key=lambda name: (name == NO_PRODUCTION, name)):
        sheet.merge_cells(start_row=line, start_column=1, end_row=line, end_column=4)
        sheet.cell(line, 1, production).font = bold
        line += 1

        people = came = 0
        for row in sorted(
            grouped[production], key=lambda item: _by_number(item["department"])
        ):
            line = _share_row(
                sheet,
                line,
                padded_number(row["department"]),
                row["people"],
                row["came"],
            )
            people += row["people"]
            came += row["came"]

        line = _share_row(sheet, line, "Итого", people, came, bold=bold)
        total_people += people
        total_came += came

    _share_row(sheet, line + 1, "Всего", total_people, total_came, bold=bold)
    return book


PRODUCTION_METHOD_WIDTHS = (16, 10, 10, 12, 10, 16, 20, 12)


def _method_share_row(sheet, line, label, people, deg, uik, uvz, u19, bold=None):
    chosen = deg + uik + uvz + u19
    cells = (
        sheet.cell(line, 1, label),
        sheet.cell(line, 2, people),
        sheet.cell(line, 3, deg),
        sheet.cell(line, 4, uik),
        sheet.cell(line, 5, uvz),
        sheet.cell(line, 6, u19),
        sheet.cell(line, 7, chosen),
        sheet.cell(line, 8, chosen / people if people else 0),
    )
    for cell in cells[1:]:
        cell.alignment = Alignment(horizontal="center")
    cells[7].number_format = "0.00%"
    if bold:
        for cell in cells:
            cell.font = bold
    return line + 1


def production_method_table():
    grouped = {}
    for row in (
        Employee.objects.exclude(department="")
        .values("production", "department")
        .annotate(
            people=Count("id"),
            deg=Count("id", filter=Q(method=DEG)),
            uik=Count("id", filter=Q(method=UIK)),
            uvz=Count("id", filter=Q(method=UVZ)),
            u19=Count("id", filter=Q(method=UIK19)),
        )
        .order_by()
    ):
        grouped.setdefault(row["production"] or NO_PRODUCTION, []).append(row)

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Способы по производствам"
    for index, width in enumerate(PRODUCTION_METHOD_WIDTHS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    bold = Font(bold=True)
    sheet.merge_cells("A1:A2")
    sheet.merge_cells("B1:B2")
    sheet.merge_cells("C1:F1")
    sheet.merge_cells("G1:H1")
    for coordinate, title in (
        ("A1", "Подразделение"),
        ("B1", "Общее количество"),
        ("C1", "Способ голосования"),
        ("C2", "ДЭГ"),
        ("D2", "УИК"),
        ("E2", "УИК-УВЗ"),
        ("F2", "УИК-19"),
        ("G1", "Итог"),
        ("G2", "Кол-во зарегестрированных"),
        ("H2", "%"),
    ):
        cell = sheet[coordinate]
        cell.value = title
        cell.font = bold
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    line = 3
    totals = dict.fromkeys(("people", "deg", "uik", "uvz", "u19"), 0)
    for production in sorted(grouped, key=lambda name: (name == NO_PRODUCTION, name)):
        sheet.merge_cells(start_row=line, start_column=1, end_row=line, end_column=7)
        sheet.cell(line, 1, production).font = bold
        line += 1

        sub = dict.fromkeys(("people", "deg", "uik", "uvz", "u19"), 0)
        for row in sorted(
            grouped[production], key=lambda item: _by_number(item["department"])
        ):
            line = _method_share_row(
                sheet,
                line,
                padded_number(row["department"]),
                row["people"],
                row["deg"],
                row["uik"],
                row["uvz"],
                row["u19"],
            )
            for key in sub:
                sub[key] += row[key]

        line = _method_share_row(
            sheet,
            line,
            "Итого",
            sub["people"],
            sub["deg"],
            sub["uik"],
            sub["uvz"],
            sub["u19"],
            bold=bold,
        )
        for key in sub:
            totals[key] += sub[key]

    _method_share_row(
        sheet,
        line + 1,
        "Всего",
        totals["people"],
        totals["deg"],
        totals["uik"],
        totals["uvz"],
        totals["u19"],
        bold=bold,
    )
    return book


def summary_table(group_field="department", group_title="Подразделение"):
    rows = sorted(
        Employee.objects.exclude(**{group_field: ""})
        .values(group_field)
        .annotate(
            people=Count("id"),
            plan_deg=Count("id", filter=Q(method=DEG)),
            plan_uik=Count("id", filter=Q(method=UIK)),
            plan_uvz=Count("id", filter=Q(method=UVZ)),
            plan_u19=Count("id", filter=Q(method=UIK19)),
            came=Count("id", filter=Q(voted=True)),
        )
        .order_by(),
        key=lambda row: _by_number(row[group_field]),
    )

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводка"
    for index, width in enumerate((18, 16, 10, 10, 12, 16, 10), 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    bold = Font(bold=True)
    centered = Alignment(horizontal="center", wrap_text=True)
    headers = (
        group_title,
        "Количество людей",
        "ДЭГ",
        "УИК",
        "УИК-УВЗ",
        "УИК-19",
        "Проголосовавшие",
        "Процент",
    )
    for column, name in enumerate(headers, 1):
        cell = sheet.cell(1, column, name)
        cell.font = bold
        cell.alignment = centered

    totals = dict.fromkeys(
        ("people", "plan_deg", "plan_uik", "plan_uvz", "plan_u19", "came"), 0
    )
    line = 2
    for row in rows:
        values = (
            row["people"],
            row["plan_deg"],
            row["plan_uik"],
            row["plan_uvz"],
            row["plan_u19"],
            row["came"],
        )
        label = row[group_field]
        if group_field == "department":
            label = padded_number(label)
        sheet.cell(line, 1, label)
        for shift, value in enumerate(values, 2):
            sheet.cell(line, shift, value).alignment = Alignment(horizontal="center")
        share = sheet.cell(line, 8, row["came"] / row["people"] if row["people"] else 0)
        share.number_format = "0.00%"
        share.alignment = Alignment(horizontal="center")
        for key in totals:
            totals[key] += row[key]
        line += 1

    sheet.cell(line, 1, "Итого").font = bold
    for shift, key in enumerate(
        ("people", "plan_deg", "plan_uik", "plan_uvz", "plan_u19", "came"), 2
    ):
        cell = sheet.cell(line, shift, totals[key])
        cell.font = bold
        cell.alignment = Alignment(horizontal="center")
    share = sheet.cell(
        line, 8, totals["came"] / totals["people"] if totals["people"] else 0
    )
    share.font = bold
    share.number_format = "0.00%"
    share.alignment = Alignment(horizontal="center")

    return book


def summary_table_no_u19(group_field="department", group_title="Подразделение"):
    rows = sorted(
        Employee.objects.exclude(**{group_field: ""})
        .exclude(okrug="19")
        .values(group_field)
        .annotate(
            people=Count("id"),
            plan_deg=Count("id", filter=Q(method=DEG)),
            plan_uik=Count("id", filter=Q(method=UIK)),
            plan_uvz=Count("id", filter=Q(method=UVZ)),
            plan_u19=Count("id", filter=Q(method=UIK19)),
            came=Count("id", filter=Q(voted=True)),
        )
        .order_by(),
        key=lambda row: _by_number(row[group_field]),
    )

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводка"
    for index, width in enumerate(SUMMARY_WIDTHS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    bold = Font(bold=True)
    centered = Alignment(horizontal="center", wrap_text=True)
    headers = (
        group_title,
        "Количество людей",
        "ДЭГ",
        "УИК",
        "УИК-УВЗ",
        "УИК-19",
        "Проголосовавшие",
        "Процент",
    )
    for column, name in enumerate(headers, 1):
        cell = sheet.cell(1, column, name)
        cell.font = bold
        cell.alignment = centered

    totals = dict.fromkeys(
        ("people", "plan_deg", "plan_uik", "plan_uvz", "plan_u19", "came"), 0
    )
    line = 2
    for row in rows:
        values = (
            row["people"],
            row["plan_deg"],
            row["plan_uik"],
            row["plan_uvz"],
            row["plan_u19"],
            row["came"],
        )
        label = row[group_field]
        if group_field == "department":
            label = padded_number(label)
        sheet.cell(line, 1, label)
        for shift, value in enumerate(values, 2):
            sheet.cell(line, shift, value).alignment = Alignment(horizontal="center")
        share = sheet.cell(line, 8, row["came"] / row["people"] if row["people"] else 0)
        share.number_format = "0.00%"
        share.alignment = Alignment(horizontal="center")
        for key in totals:
            totals[key] += row[key]
        line += 1

    sheet.cell(line, 1, "Итого").font = bold
    for shift, key in enumerate(
        ("people", "plan_deg", "plan_uik", "plan_uvz", "plan_u19", "came"), 2
    ):
        cell = sheet.cell(line, shift, totals[key])
        cell.font = bold
        cell.alignment = Alignment(horizontal="center")
    share = sheet.cell(
        line, 8, totals["came"] / totals["people"] if totals["people"] else 0
    )
    share.font = bold
    share.number_format = "0.00%"
    share.alignment = Alignment(horizontal="center")

    return book


def export_xlsx():
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сотрудники"
    sheet.append(list(COLUMNS) + ["Способ (план)", "Проголосовал", "Где голосовал"])
    fields = list(COLUMNS.values())
    for person in Employee.objects.all().iterator(chunk_size=2000):
        row = [getattr(person, f) for f in fields]
        row.append(METHOD_LABELS.get(person.method, ""))
        row.append("да" if person.voted else "нет")
        row.append(METHOD_LABELS.get(person.voted_method, ""))
        sheet.append(row)
    return book
