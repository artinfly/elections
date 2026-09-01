from datetime import datetime
from typing import Any, Optional

import openpyxl
from django.db.models import Count, Q
from django.utils import timezone
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from utils.archiver import ReportArchiver

from .helpers import (
    NO_PRODUCTION,
    REPORT_MODES,
    _apply_border,
    _by_number,
    _format_with_percent,
    department_file_name,
    padded_number,
)
from .models import DEG, UIK, UIK19, UVZ, Employee


def _share_row(
    sheet: Any, line: int, label: str, people: int, came: int, bold: Any = None
) -> int:
    """Пишет строку отчёта по производствам."""
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


def production_table() -> Any:
    """
    Отчёт "Разделение по производствам (Голосование)":
    цеха сгруппированы по производствам, для каждого — всего/проголосовало/процент.
    """
    grouped = {}
    for row in (
        Employee.objects.exclude(department="")
        .values("service", "department")
        .annotate(people=Count("id"), came=Count("id", filter=Q(voted=True)))
        .order_by()
    ):
        grouped.setdefault(row["service"] or NO_PRODUCTION, []).append(row)

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "По производствам"
    for index, width in enumerate((16, 14, 16, 10), 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A3"

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


def _method_share_row(
    sheet: Any,
    line: int,
    label: str,
    people: int,
    deg: int,
    uik: int,
    uvz: int,
    u19: int,
    bold: Any = None,
) -> int:
    """Пишет строку отчёта по способам голосования."""
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


def production_method_table(exclude_u19: bool = False) -> Any:
    """
    Отчёт "Способы голосования по производствам".
    Если exclude_u19=True, из выборки исключаются сотрудники 19 округа.
    """
    grouped = {}
    base_qs = Employee.objects.exclude(department="")
    if exclude_u19:
        base_qs = base_qs.exclude(okrug="19")

    for row in (
        base_qs.values("service", "department")
        .annotate(
            people=Count("id"),
            deg=Count("id", filter=Q(method=DEG)),
            uik=Count("id", filter=Q(method=UIK)),
            uvz=Count("id", filter=Q(method=UVZ)),
            u19=Count("id", filter=Q(method=UIK19)),
        )
        .order_by()
    ):
        grouped.setdefault(row["service"] or NO_PRODUCTION, []).append(row)

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = (
        "Способы по производствам" if not exclude_u19 else "Способы (без 19)"
    )[:31]
    for index, width in enumerate((16, 10, 10, 12, 10, 16, 20, 12), 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A3"

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


def summary_table(
    group_field: str = "department",
    group_title: str = "Подразделение",
    exclude_u19: bool = False,
) -> Any:
    """
    Сводная таблица по цехам: строка на цех — всего людей, планы по способам,
    проголосовавшие и процент явки. Внизу строка "Итого".
    Если exclude_u19=True, исключаются сотрудники 19 округа.
    """
    base_qs = Employee.objects.exclude(**{group_field: ""})
    if exclude_u19:
        base_qs = base_qs.exclude(okrug="19")

    rows = sorted(
        base_qs.values(group_field)
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

    sheet.freeze_panes = "A2"

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
        label = (
            padded_number(row[group_field])
            if group_field == "department"
            else row[group_field]
        )
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


def export_xlsx() -> Any:
    """Полная выгрузка всех сотрудников в Excel."""
    from .helpers import COLUMNS, METHOD_LABELS

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сотрудники"
    sheet.append(list(COLUMNS) + ["Способ (план)", "Проголосовал", "Где голосовал"])
    fields = list(COLUMNS.values())

    for person in (
        Employee.objects.all()
        .order_by("department", "surname", "name", "patronymic")
        .iterator(chunk_size=2000)
    ):
        row = [getattr(person, f) for f in fields]
        row.append(METHOD_LABELS.get(person.method, ""))
        row.append("да" if person.voted else "нет")
        row.append(METHOD_LABELS.get(person.voted_method, ""))
        sheet.append(row)
    return book


def department_report(
    department: str, moment: Optional[datetime] = None, mode: str = "turnout"
) -> Any:
    """Формирует Excel-отчёт по одному цеху."""
    from .helpers import REPORT_WIDTHS

    rule = REPORT_MODES[mode]
    moment = moment or timezone.localtime()
    people = Employee.objects.filter(department=department)
    total = people.count()
    done = people.filter(rule["done"]).count()

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = f"Цех {department}"[:31]
    for index, width in enumerate(REPORT_WIDTHS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A6"

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


def reports_archive(
    moment: Optional[datetime] = None, mode: str = "turnout"
) -> ReportArchiver:
    """Собирает ZIP-архив из отчётов по каждому цеху."""
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


def summary_table_no_u19(
    group_field: str = "department", group_title: str = "Подразделение"
) -> Any:
    """Обёртка для обратной совместимости — сводная таблица без сотрудников 19 округа."""
    return summary_table(
        group_field=group_field, group_title=group_title, exclude_u19=True
    )
