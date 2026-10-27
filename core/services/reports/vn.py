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

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import openpyxl
from django.utils import timezone
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ...models import DEG, UIK, UIK19, UVZ, Employee
from ..excel import _apply_border, _by_number, padded_number
from .custom import _custom_qs, _okrug19_total, _okrug19_votes
from .production import (
    BOLD,
    CENTER,
    CENTERED,
    FIRST_DATA_ROW,
    HEADER_ROW,
    draw_lead_headers,
    group_people,
    lead_headers,
    production_order,
    write_production_heading,
    write_title,
)

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


@dataclass
class _VnStats:
    """Показатели строки ВН: люди, проголосовавшие по группам, всего и 19 округ."""

    people: int
    values: list[int]
    voted: int
    total19: int
    votes19: int

    @classmethod
    def empty(cls) -> _VnStats:
        """Нулевые показатели — начальное значение для сумм."""
        return cls(0, [0] * len(VN_GROUPS), 0, 0, 0)

    @classmethod
    def of(cls, persons: list) -> _VnStats:
        """Считает показатели по списку людей."""
        return cls(
            people=len(persons),
            values=[sum(1 for p in persons if pred(p)) for _, pred in VN_GROUPS],
            voted=sum(1 for p in persons if p.voted),
            total19=sum(1 for p in persons if _okrug19_total(p)),
            votes19=sum(1 for p in persons if _okrug19_votes(p)),
        )

    def add(self, other: _VnStats) -> None:
        """Прибавляет показатели другой строки (для «Итого» и «Всего»)."""
        self.people += other.people
        self.values = [a + b for a, b in zip(self.values, other.values, strict=True)]
        self.voted += other.voted
        self.total19 += other.total19
        self.votes19 += other.votes19


@dataclass
class _VnLayout:
    """Номера колонок отчёта ВН."""

    count_col: int  # левое «Всего» (людей)
    total19_col: int  # «Всего 19 округ»
    first_group_col: int  # первая группа способов
    voted_col: int  # правое «Всего» (проголосовало)
    percent_col: int  # «Процент»
    votes19_col: int  # «Голоса Всего 19 округ»
    percent19_col: int  # «Процент 19 Округ»

    @classmethod
    def for_leads(cls, lead_count: int) -> _VnLayout:
        """Колонки идут подряд: ведущие, группы способов, хвостовые."""
        voted_col = lead_count + len(VN_GROUPS) + 1
        return cls(
            count_col=lead_count - 1,
            total19_col=lead_count,
            first_group_col=lead_count + 1,
            voted_col=voted_col,
            percent_col=voted_col + 1,
            votes19_col=voted_col + 2,
            percent19_col=voted_col + 3,
        )


def _write_vn_row(
    sheet: Any, row: int, layout: _VnLayout, stats: _VnStats, is_total_row: bool = False
) -> None:
    """
    Пишет числа строки ВН.

    В строках цехов и производств нули в группах не печатаются; в итоговых
    строках печатаются все числа, жирным. Процент — только в крайних колонках.
    """
    align = CENTERED if is_total_row else CENTER
    cell = sheet.cell(row, layout.count_col, stats.people)
    cell.alignment = align
    for offset, val in enumerate(stats.values):
        if is_total_row or val > 0:
            cell = sheet.cell(row, layout.first_group_col + offset, val)
            cell.alignment = align
            if is_total_row:
                cell.font = BOLD
    for col, value in (
        (layout.voted_col, stats.voted),
        (layout.percent_col, _vn_percent(stats.voted, stats.people)),
    ):
        cell = sheet.cell(row, col, value)
        cell.alignment = align
        if is_total_row:
            cell.font = BOLD
    if is_total_row:
        sheet.cell(row, layout.count_col).font = BOLD

    # Три колонки 19 округа: в итоговых строках жирным, выравнивание без переноса.
    for col, value in (
        (layout.total19_col, stats.total19),
        (layout.votes19_col, stats.votes19),
        (layout.percent19_col, _vn_percent(stats.votes19, stats.total19)),
    ):
        cell = sheet.cell(row, col, value)
        cell.alignment = CENTER
        if is_total_row:
            cell.font = BOLD


def _draw_vn_header(sheet: Any, lead_titles: tuple[str, ...]) -> None:
    """
    Отрисовывает двухстрочную шапку отчёта ВН (строки 2 и 3).

    Описание:
        Ведущие и хвостовые колонки (VN_TAIL_TITLES) объединяются по вертикали
        на две строки. У групп способов заголовок группы — в строке 2,
        подколонка "Проголосовал (QR-код)" — в строке 3.
    """
    column = draw_lead_headers(sheet, lead_titles)

    for group_title, _predicate in VN_GROUPS:
        for row, text in ((HEADER_ROW, group_title), (HEADER_ROW + 1, VN_SUB_TITLE)):
            cell = sheet.cell(row, column, text)
            cell.font = BOLD
            cell.alignment = CENTERED
        column += 1

    for title_text in VN_TAIL_TITLES:
        sheet.merge_cells(
            start_row=HEADER_ROW,
            start_column=column,
            end_row=HEADER_ROW + 1,
            end_column=column,
        )
        cell = sheet.cell(HEADER_ROW, column, title_text)
        cell.font = BOLD
        cell.alignment = CENTERED
        column += 1


# ==============================================================================
# Отчёт
# ==============================================================================


def custom_production_summary_vn(
    params: dict,
    include_depts: bool = False,
    moment: datetime | None = None,
) -> Workbook:
    """
    Генерирует сводный отчёт по производствам в формате ВН.

    Описание:
        Агрегирует данные по производствам (поле service), опционально
        с разбивкой по цехам. Сотрудники с отметкой "Уважительная причина",
        которые не проголосовали, исключаются из выборки целиком: они не
        попадают в левое "Всего" и не учитываются в процентах.

        Структура колонок:
          с цехами:  Производство | Цех | Всего | Всего 19 округ | 4 группы (QR)
                     | Всего | Процент | Голоса Всего 19 округ | Процент 19 Округ
          без цехов: Производство | Всего | Всего 19 округ | 4 группы (QR)
                     | Всего | Процент | Голоса Всего 19 округ | Процент 19 Округ

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
    people = (
        _custom_qs(params)
        .exclude(absence=True, voted=False)
        .order_by("service", "department", "surname", "name", "patronymic")
    )

    leads = lead_headers(include_depts)
    layout = _VnLayout.for_leads(len(leads))
    total_columns = len(leads) + len(VN_GROUPS) + len(VN_TAIL_TITLES)

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводный отчёт"

    write_title(sheet, moment, total_columns)
    _draw_vn_header(sheet, leads)

    lead_widths = (16, 14, 10, 12) if include_depts else (24, 10, 12)
    widths = lead_widths + (14,) * len(VN_GROUPS) + (10, 12, 14, 12)
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A4"

    data = group_people(people.iterator(chunk_size=2000), include_depts)
    grand = _VnStats.empty()
    row = FIRST_DATA_ROW

    for production in production_order(data):
        if not include_depts:
            # Одна строка на производство.
            stats = _VnStats.of(data[production])
            sheet.cell(row, 1, production)
            _write_vn_row(sheet, row, layout, stats)
            grand.add(stats)
            row += 1
            continue

        # Название производства, строки его цехов и строка «Итого».
        write_production_heading(sheet, row, production, total_columns)
        row += 1
        subtotal = _VnStats.empty()
        for department in sorted(data[production], key=_by_number):
            stats = _VnStats.of(data[production][department])
            sheet.cell(row, 1, production)
            sheet.cell(row, 2, padded_number(department))
            _write_vn_row(sheet, row, layout, stats)
            subtotal.add(stats)
            row += 1
        sheet.cell(row, layout.count_col - 1, "Итого").font = BOLD
        _write_vn_row(sheet, row, layout, subtotal, is_total_row=True)
        grand.add(subtotal)
        row += 1

    sheet.cell(row, layout.count_col - 1, "Всего по Обществу").font = BOLD
    _write_vn_row(sheet, row, layout, grand, is_total_row=True)
    _apply_border(sheet)
    return book
