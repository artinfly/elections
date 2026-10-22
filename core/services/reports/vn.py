"""
Модуль генерации сводных отчётов по производствам в формате ВН.

Описание:
    Два новых режима группировки конструктора отчётов:
      - "По производствам (с цехами, ВН)";
      - "По производствам (без цехов, ВН)".

    Отличия формата ВН от компактного:
      - в каждой группе способов остаётся только подколонка
        "Проголосовал (QR-код)";
      - группа "УИК-19" переименована в "Открепленные";
      - проценты НЕ пишутся в ячейки "Всего" и "Итого", а вынесены
        в отдельный крайний столбец "Процент" (правое "Всего" / левое "Всего"),
        с точностью до двух знаков после запятой;
      - колонки плана нет ни в одном из двух отчётов;
      - добавлены колонки 19 округа: "Всего 19 округ" (после левого "Всего")
        = "Всего людей" минус количество в выборке по условию «не 19 округ»;
        после колонки "Процент" — "Голоса Всего 19 округ" = проголосовавшие
        минус проголосовавшие в выборке по тому же условию с voted_method вместо
        method, и "Процент 19 Округ";
      - сотрудники с отметкой "Уважительная причина" (absence) не участвуют
        в отчёте: они не попадают в левое "Всего" и не влияют на процент.
"""

from collections import defaultdict
from datetime import datetime
from typing import Callable, Optional

import openpyxl
from django.utils import timezone
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ...models import DEG, UIK, UIK19, UVZ, Employee
from ..excel import NO_PRODUCTION, _apply_border, _by_number, padded_number
from .custom import _custom_qs, _okrug19_total, _okrug19_votes

# ==============================================================================
# Конфигурация отчёта ВН
# ==============================================================================

# Единственная подколонка во всех группах способов.
VN_SUB_TITLE = "Проголосовал (QR-код)"

# Группы отчёта ВН: (заголовок группы, предикат отметки).
# Четвёртая группа — переименованный "УИК-19": считаем проголосовавших QR по у19.
VN_GROUPS: list[tuple[str, Callable[[Employee], bool]]] = [
    ("ДЭГ", lambda p: p.voted and p.voted_method == DEG),
    ("На участке", lambda p: p.voted and p.voted_method == UIK),
    ("На участке УВЗ", lambda p: p.voted and p.voted_method == UVZ),
    ("Открепленные", lambda p: p.voted and p.voted_method == UIK19),
]


# Заголовки завершающих колонок отчёта (после групп способов).
VN_TAIL_TITLES = ("Всего", "Процент", "Голоса Всего 19 округ", "Процент 19 Округ")


def _vn_percent(voted: int, total: int) -> str:
    """
    Процент проголосовавших от общего числа людей в строке.

    Описание:
        Считается как правое "Всего" (проголосовало) / левое "Всего" (людей).
        Формат — точное значение с двумя знаками после запятой, запятая
        как разделитель (например, "74,70%"), как в макете.
        При нулевом знаменателе возвращает пустую строку, чтобы не делить на ноль.

    Аргументы:
        voted: количество проголосовавших (правое "Всего").
        total: количество людей в строке (левое "Всего").

    Возвращает:
        str: строка процента или пустая строка.
    """
    if not total:
        return ""
    return f"{voted / total * 100:.2f}".replace(".", ",") + "%"


def _vn_row_stats(persons: list) -> tuple[list[int], int, int, int, int]:
    """
    Считает показатели одной строки отчёта (цеха, производства или итога).

    Возвращает:
        tuple: (список значений по группам VN_GROUPS,
                всего людей,
                всего проголосовало,
                всего 19 округ,
                голоса всего 19 округ).
    """
    values = [sum(1 for p in persons if predicate(p)) for _, predicate in VN_GROUPS]
    voted = sum(1 for p in persons if p.voted)
    total19 = sum(1 for p in persons if _okrug19_total(p))
    votes19 = sum(1 for p in persons if _okrug19_votes(p))
    return values, len(persons), voted, total19, votes19


def _write_vn_okrug19(
    sheet,
    row: int,
    total19_col: int,
    votes19_col: int,
    percent19_col: int,
    total19: int,
    votes19: int,
    is_total_row: bool = False,
) -> None:
    """
    Пишет в строку три колонки 19 округа: "Всего 19 округ", "Голоса Всего
    19 округ" и "Процент 19 Округ" (голоса / всего, формат как у "Процент").
    """
    values = (
        (total19_col, total19),
        (votes19_col, votes19),
        (percent19_col, _vn_percent(votes19, total19)),
    )
    for col, value in values:
        cell = sheet.cell(row, col, value)
        cell.alignment = Alignment(horizontal="center")
        if is_total_row:
            cell.font = Font(bold=True)


def _draw_vn_header(sheet, lead_headers: tuple[str, ...]) -> None:
    """
    Отрисовывает двухстрочную шапку отчёта ВН (строки 2 и 3).

    Описание:
        Ведущие колонки и хвостовые (VN_TAIL_TITLES) объединяются
        по вертикали на две строки. Группы способов рисуются так:
        заголовок группы в строке 2, подколонка "Проголосовал (QR-код)"
        в строке 3.
    """
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    column = 1
    for title_text in lead_headers:
        sheet.merge_cells(
            start_row=2, start_column=column, end_row=3, end_column=column
        )
        cell = sheet.cell(2, column, title_text)
        cell.font = bold
        cell.alignment = centered
        column += 1

    for group_title, _predicate in VN_GROUPS:
        head = sheet.cell(2, column, group_title)
        head.font = bold
        head.alignment = centered
        sub = sheet.cell(3, column, VN_SUB_TITLE)
        sub.font = bold
        sub.alignment = centered
        column += 1

    for title_text in VN_TAIL_TITLES:
        sheet.merge_cells(
            start_row=2, start_column=column, end_row=3, end_column=column
        )
        cell = sheet.cell(2, column, title_text)
        cell.font = bold
        cell.alignment = centered
        column += 1


# ==============================================================================
# Отчёт
# ==============================================================================


def custom_production_summary_vn(
    params: dict,
    include_depts: bool = False,
    moment: Optional[datetime] = None,
) -> Workbook:
    """
    Генерирует сводный отчёт по производствам в формате ВН.

    Описание:
        Агрегирует данные по производствам (поле service), опционально
        с разбивкой по цехам. Сотрудники с отметкой "Уважительная причина"
        (absence=True) исключаются из выборки целиком: они не попадают
        в левое "Всего" и не учитываются в процентах.

        Структура колонок:
          с цехами:  Производство | Цех | Всего | Всего 19 округ | 4 группы (QR)
                     | Всего | Процент | Голоса Всего 19 округ | Процент 19 Округ
          без цехов: Производство | Всего | Всего 19 округ | 4 группы (QR)
                     | Всего | Процент | Голоса Всего 19 округ | Процент 19 Округ

        В детальных строках нули в группах не печатаются (пустая ячейка),
        в строках "Итого" и "Всего по Обществу" печатаются все числа,
        но без процентов внутри ячеек — процент только в крайнем столбце
        и всегда с двумя знаками после запятой.

    Аргументы:
        params: параметры фильтров конструктора.
        include_depts: если True, выводит строки по цехам внутри производств
            и строки "Итого" по каждому производству.
            Если False — одна строка на производство.
        moment: момент времени для заголовка отчёта
            (по умолчанию — текущее время на ПК).

    Возвращает:
        Workbook: книга openpyxl с готовым отчётом.
    """
    moment = moment or timezone.localtime()

    # Люди с уважительной причиной не участвуют в отчёте ВН.
    people = (
        _custom_qs(params)
        .exclude(absence=True, voted=False)
        .order_by("service", "department", "surname", "name", "patronymic")
    )

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводный отчёт"
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    lead_headers = (
        ("Производство", "Цех", "Всего", "Всего 19 округ")
        if include_depts
        else ("Производство", "Всего", "Всего 19 округ")
    )
    total_columns = len(lead_headers) + len(VN_GROUPS) + len(VN_TAIL_TITLES)

    # Главный заголовок отчёта (строка 1).
    title = sheet.cell(
        1, 1, f"Итоговый отчёт по видам производства на {moment:%d.%m.%y %H:%M}"
    )
    title.font = Font(bold=True, size=12)
    title.alignment = centered
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_columns)

    # Шапка колонок (строки 2-3).
    _draw_vn_header(sheet, lead_headers)

    # Ширины колонок.
    widths = (
        ((16, 14, 10, 12) if include_depts else (24, 10, 12))
        + (14,) * len(VN_GROUPS)
        + (10, 12, 14, 12)
    )
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    sheet.freeze_panes = "A4"

    # Индексы колонок: "Всего 19 округ" (последняя ведущая), группы,
    # правое "Всего", процент, голоса 19 округа, процент 19 округа.
    lead_count = len(lead_headers)
    total19_col = lead_count
    first_group_col = lead_count + 1
    total_voted_col = lead_count + len(VN_GROUPS) + 1
    percent_col = total_voted_col + 1
    votes19_col = percent_col + 1
    percent19_col = votes19_col + 1

    # Группировка данных в памяти (как в текущих отчётах — по полю service).
    if include_depts:
        data = defaultdict(lambda: defaultdict(list))
    else:
        data = defaultdict(list)

    for person in people.iterator(chunk_size=2000):
        production_name = person.service or NO_PRODUCTION
        if include_depts:
            data[production_name][person.department].append(person)
        else:
            data[production_name].append(person)

    row = 4
    grand_total_people = 0
    grand_totals = [0] * len(VN_GROUPS)
    grand_voted = 0
    grand_total19 = 0
    grand_votes19 = 0

    # Сортировка производств: обычные по алфавиту, "Без производства" в конце.
    for production in sorted(
        data.keys(), key=lambda name: (name == NO_PRODUCTION, name)
    ):
        if include_depts:
            # Заголовок производства на всю ширину таблицы.
            sheet.merge_cells(
                start_row=row, start_column=1, end_row=row, end_column=total_columns
            )
            sheet.cell(row, 1, production).font = bold
            sheet.cell(row, 1).alignment = Alignment(horizontal="center")
            row += 1

            prod_total_people = 0
            prod_totals = [0] * len(VN_GROUPS)
            prod_voted = 0
            prod_total19 = 0
            prod_votes19 = 0

            # Сортировка цехов внутри производства.
            for department in sorted(data[production].keys(), key=_by_number):
                persons = data[production][department]
                values, count, voted, total19, votes19 = _vn_row_stats(persons)
                prod_total_people += count
                prod_voted += voted
                grand_total_people += count
                grand_voted += voted
                prod_total19 += total19
                prod_votes19 += votes19
                grand_total19 += total19
                grand_votes19 += votes19

                sheet.cell(row, 1, production)
                sheet.cell(row, 2, padded_number(department))
                sheet.cell(row, 3, count).alignment = Alignment(horizontal="center")
                for offset, val in enumerate(values):
                    if val > 0:
                        sheet.cell(row, first_group_col + offset, val).alignment = (
                            Alignment(horizontal="center")
                        )
                    prod_totals[offset] += val
                    grand_totals[offset] += val
                sheet.cell(row, total_voted_col, voted).alignment = Alignment(
                    horizontal="center"
                )
                sheet.cell(row, percent_col, _vn_percent(voted, count)).alignment = (
                    Alignment(horizontal="center")
                )
                _write_vn_okrug19(
                    sheet,
                    row,
                    total19_col,
                    votes19_col,
                    percent19_col,
                    total19,
                    votes19,
                )
                row += 1

            # Строка "Итого" по производству: только числа, без процентов в ячейках.
            sheet.cell(row, 2, "Итого").font = bold
            sheet.cell(row, 3, prod_total_people).font = bold
            sheet.cell(row, 3).alignment = centered
            for offset, val in enumerate(prod_totals):
                cell = sheet.cell(row, first_group_col + offset, val)
                cell.font = bold
                cell.alignment = centered
            cell = sheet.cell(row, total_voted_col, prod_voted)
            cell.font = bold
            cell.alignment = centered
            cell = sheet.cell(
                row, percent_col, _vn_percent(prod_voted, prod_total_people)
            )
            cell.font = bold
            cell.alignment = centered
            _write_vn_okrug19(
                sheet,
                row,
                total19_col,
                votes19_col,
                percent19_col,
                prod_total19,
                prod_votes19,
                is_total_row=True,
            )
            row += 1
        else:
            # Режим без разбивки по цехам: строка на производство.
            persons = data[production]
            values, count, voted, total19, votes19 = _vn_row_stats(persons)
            grand_total_people += count
            grand_voted += voted
            grand_total19 += total19
            grand_votes19 += votes19

            sheet.cell(row, 1, production)
            sheet.cell(row, 2, count).alignment = Alignment(horizontal="center")
            for offset, val in enumerate(values):
                if val > 0:
                    sheet.cell(row, first_group_col + offset, val).alignment = (
                        Alignment(horizontal="center")
                    )
                grand_totals[offset] += val
            sheet.cell(row, total_voted_col, voted).alignment = Alignment(
                horizontal="center"
            )
            sheet.cell(row, percent_col, _vn_percent(voted, count)).alignment = (
                Alignment(horizontal="center")
            )
            _write_vn_okrug19(
                sheet, row, total19_col, votes19_col, percent19_col, total19, votes19
            )
            row += 1

    # Строка "Всего по Обществу".
    if include_depts:
        sheet.cell(row, 2, "Всего по Обществу").font = bold
        sheet.cell(row, 3, grand_total_people).font = bold
        sheet.cell(row, 3).alignment = centered
        for offset, val in enumerate(grand_totals):
            cell = sheet.cell(row, first_group_col + offset, val)
            cell.font = bold
            cell.alignment = centered
        cell = sheet.cell(row, total_voted_col, grand_voted)
        cell.font = bold
        cell.alignment = centered
        cell = sheet.cell(
            row, percent_col, _vn_percent(grand_voted, grand_total_people)
        )
        cell.font = bold
        cell.alignment = centered
        _write_vn_okrug19(
            sheet,
            row,
            total19_col,
            votes19_col,
            percent19_col,
            grand_total19,
            grand_votes19,
            is_total_row=True,
        )
    else:
        sheet.cell(row, 1, "Всего по Обществу").font = bold
        sheet.cell(row, 2, grand_total_people).font = bold
        sheet.cell(row, 2).alignment = centered
        for offset, val in enumerate(grand_totals):
            cell = sheet.cell(row, first_group_col + offset, val)
            cell.font = bold
            cell.alignment = centered
        cell = sheet.cell(row, total_voted_col, grand_voted)
        cell.font = bold
        cell.alignment = centered
        cell = sheet.cell(
            row, percent_col, _vn_percent(grand_voted, grand_total_people)
        )
        cell.font = bold
        cell.alignment = centered
        _write_vn_okrug19(
            sheet,
            row,
            total19_col,
            votes19_col,
            percent19_col,
            grand_total19,
            grand_votes19,
            is_total_row=True,
        )

    _apply_border(sheet)
    return book
