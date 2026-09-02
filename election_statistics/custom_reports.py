"""
Кастомные (динамические) отчёты по фильтрам конструктора.

Сводный отчёт "по людям" (custom_report) и сводный по производствам
(custom_production_summary). Вся семантика галочек собрана в одном
списке CUSTOM_COLUMNS — если смысл колонки изменится, правится только он.
"""

from collections import defaultdict
from datetime import datetime
from typing import Any, Callable, Optional

import openpyxl
from django.db.models import Q
from django.utils import timezone
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .helpers import (
    NO_PRODUCTION,
    _apply_border,
    _by_number,
    _format_with_percent,
    padded_number,
)
from .models import DEG, UIK, UIK19, UVZ, Employee

# ==============================================================================
# Конфигурация колонок сводного отчёта
# ==============================================================================

# (Группа, Подколонка, Предикат-проверка). Вся бизнес-логика галочек — здесь.
# ВНИМАНИЕ (см. аудит): тесты и README описывают другую семантику группы
# "Не 19 округ" (по полю okrug, без колонок регистрации) и колонку
# "Не пойдет" вместо "Не определился". Ниже — текущая бизнес-версия;
# при возврате к okrug-семантике правится только этот список.
CUSTOM_COLUMNS: list[tuple[str, str, Callable]] = [
    ("ДЭГ", "Планирует", lambda p: p.method == DEG),
    ("ДЭГ", "Зарегистрирован", lambda p: p.mark_deg),
    ("ДЭГ", "Проголосовал", lambda p: p.voted and p.voted_method == DEG),
    ("На участке", "Планирует", lambda p: p.method == UIK),
    ("На участке", "Проголосовал", lambda p: p.voted and p.voted_method == UIK),
    ("На участке УВЗ", "Планирует", lambda p: p.method == UVZ),
    ("На участке УВЗ", "Заявление оформил", lambda p: p.mark_uvz),
    ("На участке УВЗ", "Проголосовал", lambda p: p.voted and p.voted_method == UVZ),
    # Группа считает привязанных к участку 19 округа, независимо от округа сотрудника
    ("Не 19 округ", "Планирует", lambda p: p.method == UIK19),
    ("Не 19 округ", "Открепился", lambda p: p.method == UIK19 and p.detached),
    ("Не 19 округ", "Проголосовал", lambda p: p.voted and p.voted_method == UIK19),
    ("", "Не определился", lambda p: p.method not in (DEG, UIK, UVZ, UIK19)),
]


# ==============================================================================
# Фильтрация и заголовки
# ==============================================================================


def _custom_qs(params: dict) -> Any:
    """
    Строит QuerySet сотрудников по параметрам формы конструктора.

    Аргументы:
        params: словарь параметров (production, service, dep, method,
            where, okrug включая "none" и "20+21", uik).

    Возвращает:
        Отфильтрованный QuerySet.
    """
    qs = Employee.objects.all()

    production = (params.get("production") or "").strip()
    dep = (params.get("dep") or "").strip()
    method = (params.get("method") or "").strip()
    where = (params.get("where") or "").strip()
    okrug = (params.get("okrug") or "").strip()
    uik = (params.get("uik") or "").strip()
    service = (params.get("service") or "").strip()

    if production:
        qs = qs.filter(production=production)
    if service:
        qs = qs.filter(service=service)
    if dep:
        qs = qs.filter(department=dep)

    if method == "none":
        qs = qs.filter(method="")
    elif method:
        qs = qs.filter(method=method)

    if where == "none":
        qs = qs.filter(voted_method="")
    elif where:
        qs = qs.filter(voted_method=where)

    if okrug == "none":
        qs = qs.filter(okrug="")
    elif okrug == "20+21":
        qs = qs.filter(okrug__in=["20", "21"])
    elif okrug:
        qs = qs.filter(okrug=okrug)

    if uik:
        qs = qs.filter(uik=uik)

    return qs


def _custom_groups() -> list[tuple[str, list[tuple[str, Callable]]]]:
    """
    Группирует колонки CUSTOM_COLUMNS по имени группы.

    Возвращает:
        Список (группа, [(подколонка, предикат)...]) в порядке колонок.
    """
    groups = []
    for group, sub, predicate in CUSTOM_COLUMNS:
        if groups and groups[-1][0] == group:
            groups[-1][1].append((sub, predicate))
        else:
            groups.append((group, [(sub, predicate)]))
    return groups


def _draw_custom_groups(sheet: Any, start_row: int, start_col: int) -> list[Callable]:
    """
    Отрисовывает объединённые заголовки групп и подколонки в Excel.

    Группа без имени (пустая строка) рисуется как одиночная колонка,
    объединённая по вертикали на две строки.

    Аргументы:
        sheet: лист openpyxl.
        start_row: строка, с которой начинается шапка.
        start_col: колонка, с которой начинается шапка.

    Возвращает:
        Список предикатов в порядке колонок листа.
    """
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    column = start_col
    predicates = []
    for group, subs in _custom_groups():
        start, end = column, column + len(subs) - 1
        if group:
            sheet.merge_cells(
                start_row=start_row,
                start_column=start,
                end_row=start_row,
                end_column=end,
            )
            head = sheet.cell(start_row, column, group)
        else:
            sheet.merge_cells(
                start_row=start_row,
                start_column=start,
                end_row=start_row + 1,
                end_column=end,
            )
            head = sheet.cell(start_row, column, subs[0][0])

        head.font = bold
        head.alignment = centered

        for sub, predicate in subs:
            if group:
                cell = sheet.cell(start_row + 1, column, sub)
                cell.font = bold
                cell.alignment = centered
            predicates.append(predicate)
            column += 1
    return predicates


# ==============================================================================
# Отчёты
# ==============================================================================


def custom_report(params: dict) -> Any:
    """
    Генерирует сводный Excel-отчёт "По людям" на основе фильтров.

    Аргументы:
        params: параметры фильтров конструктора.

    Возвращает:
        Книга openpyxl: строка на сотрудника, галочки по предикатам,
        внизу строка ИТОГО.
    """
    people = _custom_qs(params).order_by(
        "department", "surname", "name", "patronymic", "uik"
    )

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводный отчёт"
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Ведущие колонки
    column = 1
    for title in ("Номер УИК", "Цех", "Таб.№", "ФИО", "Округ"):
        sheet.merge_cells(
            start_row=1, start_column=column, end_row=2, end_column=column
        )
        cell = sheet.cell(1, column, title)
        cell.font = bold
        cell.alignment = centered
        column += 1

    predicates = _draw_custom_groups(sheet, start_row=1, start_col=column)

    widths = (10, 8, 10, 42, 8) + (12,) * len(predicates)
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A3"

    row = 3
    totals = [0] * len(predicates)
    total_people = 0

    # iterator с chunk_size не держит всю выборку в памяти при чтении
    for person in people.iterator(chunk_size=2000):
        total_people += 1
        sheet.cell(row, 1, person.uik)
        sheet.cell(row, 2, person.department)
        sheet.cell(row, 3, person.tab_number)
        sheet.cell(row, 4, person.fio)
        sheet.cell(row, 5, person.okrug)

        for offset, predicate in enumerate(predicates):
            if predicate(person):
                sheet.cell(row, 6 + offset, "+")
                totals[offset] += 1
        row += 1

    # Строка ИТОГО
    sheet.cell(row, 1, "ИТОГО").font = bold
    sheet.cell(row, 2, total_people).font = bold
    sheet.cell(row, 2).alignment = centered

    for offset, total in enumerate(totals):
        cell = sheet.cell(row, 6 + offset, _format_with_percent(total, total_people))
        cell.font = bold
        cell.alignment = centered

    _apply_border(sheet)
    return book


def custom_production_summary(
    params: dict, include_depts: bool = False, moment: Optional[datetime] = None
) -> Any:
    """
    Генерирует сводный отчёт с группировкой по производствам
    (опционально с разбивкой по цехам).

    Аргументы:
        params: параметры фильтров конструктора.
        include_depts: выводить строки по цехам внутри производства.
        moment: момент для заголовка (по умолчанию текущее время).

    Возвращает:
        Книга openpyxl со сводным отчётом.
    """
    moment = moment or timezone.localtime()
    # ВНИМАНИЕ (см. аудит): группировка по полю production, а стандартные
    # отчёты "по производствам" идут по service. Привести к единому полю.
    people = _custom_qs(params).order_by(
        "production", "department", "surname", "name", "patronymic"
    )

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводный отчёт"
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    predicates = []
    for group, subs in _custom_groups():
        for sub, predicate in subs:
            predicates.append(predicate)

    lead_headers = (
        ("Производство", "Цех", "Всего") if include_depts else ("Производство", "Всего")
    )
    total_columns = len(lead_headers) + len(predicates)

    title = sheet.cell(
        1, 1, f"Итоговый отчёт по видам производства на {moment:%d.%m.%y %H:%M}"
    )
    title.font = Font(bold=True, size=12)
    title.alignment = centered
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_columns)

    header_row = 2
    column = 1
    for title_text in lead_headers:
        sheet.merge_cells(
            start_row=header_row,
            start_column=column,
            end_row=header_row + 1,
            end_column=column,
        )
        cell = sheet.cell(header_row, column, title_text)
        cell.font = bold
        cell.alignment = centered
        column += 1

    _draw_custom_groups(sheet, start_row=header_row, start_col=column)

    widths = (
        (16, 14, 10) + (12,) * len(predicates)
        if include_depts
        else (24, 10) + (12,) * len(predicates)
    )
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A4"

    # Группировка в памяти: на этих объёмах проще, чем повторные запросы
    if include_depts:
        data = defaultdict(lambda: defaultdict(list))
    else:
        data = defaultdict(list)

    for person in people.iterator(chunk_size=2000):
        production = person.production or NO_PRODUCTION
        if include_depts:
            data[production][person.department].append(person)
        else:
            data[production].append(person)

    row = header_row + 2
    grand_total_people = 0
    grand_totals = [0] * len(predicates)

    # "Без производства" всегда последняя группа
    for production in sorted(
        data.keys(), key=lambda name: (name == NO_PRODUCTION, name)
    ):
        if include_depts:
            sheet.merge_cells(
                start_row=row, start_column=1, end_row=row, end_column=total_columns
            )
            sheet.cell(row, 1, production).font = bold
            sheet.cell(row, 1).alignment = Alignment(horizontal="center")
            row += 1

            prod_total_people = 0
            prod_totals = [0] * len(predicates)

            for department in sorted(data[production].keys(), key=_by_number):
                persons = data[production][department]
                count = len(persons)
                prod_total_people += count
                grand_total_people += count

                sheet.cell(row, 1, production)
                sheet.cell(row, 2, padded_number(department))
                sheet.cell(row, 3, count).alignment = Alignment(horizontal="center")

                for offset, predicate in enumerate(predicates):
                    val = sum(1 for p in persons if predicate(p))
                    if val > 0:
                        sheet.cell(row, 4 + offset, val).alignment = Alignment(
                            horizontal="center"
                        )
                    prod_totals[offset] += val
                    grand_totals[offset] += val
                row += 1

            sheet.cell(row, 2, "Итого").font = bold
            sheet.cell(row, 3, prod_total_people).font = bold
            sheet.cell(row, 3).alignment = Alignment(horizontal="center")
            for offset, val in enumerate(prod_totals):
                cell = sheet.cell(
                    row, 4 + offset, _format_with_percent(val, prod_total_people)
                )
                cell.font = bold
                cell.alignment = centered
            row += 1
        else:
            persons = data[production]
            count = len(persons)
            grand_total_people += count

            sheet.cell(row, 1, production)
            sheet.cell(row, 2, count).alignment = Alignment(horizontal="center")

            for offset, predicate in enumerate(predicates):
                val = sum(1 for p in persons if predicate(p))
                if val > 0:
                    sheet.cell(row, 3 + offset, val).alignment = Alignment(
                        horizontal="center"
                    )
                grand_totals[offset] += val
            row += 1

    if include_depts:
        sheet.cell(row, 2, "Всего по Обществу").font = bold
        sheet.cell(row, 3, grand_total_people).font = bold
        sheet.cell(row, 3).alignment = Alignment(horizontal="center")
        for offset, val in enumerate(grand_totals):
            cell = sheet.cell(
                row, 4 + offset, _format_with_percent(val, grand_total_people)
            )
            cell.font = bold
            cell.alignment = centered
    else:
        sheet.cell(row, 1, "Всего по Обществу").font = bold
        sheet.cell(row, 2, grand_total_people).font = bold
        sheet.cell(row, 2).alignment = Alignment(horizontal="center")
        for offset, val in enumerate(grand_totals):
            cell = sheet.cell(
                row, 3 + offset, _format_with_percent(val, grand_total_people)
            )
            cell.font = bold
            cell.alignment = centered

    _apply_border(sheet)
    return book
