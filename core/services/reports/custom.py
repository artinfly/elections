"""
Модуль генерации кастомных (динамических) отчётов по фильтрам конструктора.

Описание:
    Содержит логику создания сводных отчётов, настраиваемых пользователем
    через форму фильтров. Включает отчёт "по людям" (каждый сотрудник строкой)
    и отчёт "по производствам" (агрегированные данные).
    Вся бизнес-логика отображения галочек в колонках сосредоточена в списке CUSTOM_COLUMNS.
"""

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Callable, Optional

import openpyxl
from django.db.models import QuerySet
from django.utils import timezone
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from ...models import DEG, UIK, UIK19, UVZ, Employee
from ..excel import (
    _apply_border,
    _by_number,
    _format_with_percent,
    padded_number,
)
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
# Конфигурация колонок сводного отчёта
# ==============================================================================

# Список колонок отчёта в формате: (Имя группы, Имя подколонки, Предикат-проверка).
#
# Структура:
#   - Имя группы: название группы колонок (например, "ДЭГ", "На участке").
#     Пустая строка означает одиночную колонку без общего заголовка.
#   - Имя подколонки: название конкретной колонки внутри группы.
#   - Предикат: функция, принимающая объект Employee и возвращающая True,
#     если для сотрудника нужно поставить отметку "+" в этой колонке.
#
# ВАЖНО: Если бизнес-логика колонки изменится, править нужно ТОЛЬКО этот список.
# Не нужно менять функции отрисовки или подсчёта — они автоматически адаптируются.
CUSTOM_COLUMNS: list[tuple[str, str, Callable[[Employee], bool]]] = [
    # Группа ДЭГ (дистанционное электронное голосование)
    ("ДЭГ", "Планирует", lambda p: p.method == DEG),
    ("ДЭГ", "Проголосовал (QR-код)", lambda p: p.voted and p.voted_method == DEG),
    (
        "ДЭГ",
        "Проголосовал (ответственный)",
        lambda p: p.proxy_vote and p.method == DEG,
    ),
    # Группа голосования на обычном участке
    ("На участке", "Планирует", lambda p: p.method == UIK),
    (
        "На участке",
        "Проголосовал (QR-код)",
        lambda p: p.voted and p.voted_method == UIK,
    ),
    (
        "На участке",
        "Проголосовал (ответственный)",
        lambda p: p.proxy_vote and p.method == UIK,
    ),
    # Группа голосования на участке УВЗ (на предприятии)
    ("На участке УВЗ", "Планирует", lambda p: p.method == UVZ),
    ("На участке УВЗ", "Заявление оформил", lambda p: p.method == UVZ and p.mark_uvz),
    (
        "На участке УВЗ",
        "Проголосовал (QR-код)",
        lambda p: p.voted and p.voted_method == UVZ,
    ),
    (
        "На участке УВЗ",
        "Проголосовал (ответственный)",
        lambda p: p.proxy_vote and p.method == UVZ,
    ),
    # Группа 19-го округа
    ("УИК-19", "Планирует", lambda p: p.method == UIK19),
    ("УИК-19", "Проголосовал (QR-код)", lambda p: p.voted and p.voted_method == UIK19),
    (
        "УИК-19",
        "Проголосовал (ответственный)",
        lambda p: p.proxy_vote and p.method == UIK19,
    ),
    # Одиночные колонки без группы (имя группы пустое).
    # "Не определился" — сотрудник не выбрал способ.
    # "Отсутствовал по УП" — сотрудник отметил уважительную причину отсутствия.
    (
        "",
        "Не определился",
        lambda p: p.method not in (DEG, UIK, UVZ, UIK19) and not p.absence,
    ),
    ("", "Уважительная причина", lambda p: p.absence),
]


# Предикат "Планирует" для каждой группы способов голосования, полученный
# из CUSTOM_COLUMNS (без дублирования логики). Используется, чтобы считать
# процент по колонкам "Проголосовал (QR-код)" и "Проголосовал (ответственный)"
# не от общего числа людей, а от числа людей, планирующих голосовать этим
# способом — отдельно для каждой группы.
GROUP_PLAN_PREDICATE: dict[str, Callable[[Employee], bool]] = {
    group: predicate for group, sub, predicate in CUSTOM_COLUMNS if sub == "Планирует"
}


# ==============================================================================
# Колонки "19 округ" (только для отчёта по производствам)
# ==============================================================================

# Города, жители которых без указанного округа не относятся к 19 округу.
OKRUG19_EXCLUDED_CITIES = ("Нижний Тагил", "Покровское-1")

# Заголовки колонок, добавляемых ПОСЛЕ последней колонки "Всего" (процент).
OKRUG19_TAIL_HEADERS = ("Голоса Всего 19 округ", "Процент 19 Округ")


def _okrug19_base(p: Employee) -> bool:
    """
    Базовое условие «не 19 округ» (без проверки способа):

        (округ пустой и город не из OKRUG19_EXCLUDED_CITIES)
        или округ 20, или округ 21
    """
    okrug = (p.okrug or "").strip()
    return (okrug == "" and p.city not in OKRUG19_EXCLUDED_CITIES) or okrug in (
        "20",
        "21",
    )


def _okrug19_cond_total(p: Employee) -> bool:
    """Полное условие «не 19 округ»: базовое условие и план не «УИК-19»."""
    return _okrug19_base(p) and p.method != UIK19


def _okrug19_cond_votes(p: Employee) -> bool:
    """То же условие, но вместо method используется voted_method."""
    return _okrug19_base(p) and p.voted_method != UIK19


def _okrug19_total(p: Employee) -> bool:
    """
    Колонка "Всего 19 округ" = "Всего людей" минус количество людей в выборке
    по условию «не 19 округ» (сотрудник считается, если условие
    для него НЕ выполняется).
    """
    return not _okrug19_cond_total(p)


def _okrug19_votes(p: Employee) -> bool:
    """
    Колонка "Голоса Всего 19 округ" = проголосовавшие (voted) минус
    проголосовавшие, попавшие в выборку по тому же условию, но с voted_method
    вместо method (сотрудник считается, если он проголосовал и условие
    для него НЕ выполняется).

    Голос засчитывается только тем, кто входит и в "Всего 19 округ"
    (условие с method тоже НЕ выполняется). Иначе человек, у которого
    method != 'u19', а voted_method == 'u19', попадал бы в голоса, но не
    в "Всего 19 округ", и "Процент 19 Округ" мог превысить 100%.
    """
    return bool(p.voted) and _okrug19_total(p) and not _okrug19_cond_votes(p)


def _okrug19_percent(votes: int, total: int) -> str:
    """Колонка "Процент 19 Округ": голоса 19 округ / всего 19 округ."""
    return f"{round(votes / total * 100)}%" if total else ""


def _draw_okrug19_tail_headers(sheet: Any, header_row: int, column: int) -> None:
    """Рисует заголовки двух завершающих колонок 19 округа (на две строки шапки)."""
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for title_text in OKRUG19_TAIL_HEADERS:
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


def _write_okrug19_tail(
    sheet: Any,
    row: int,
    column: int,
    total: int,
    votes: int,
    is_total_row: bool = False,
) -> None:
    """
    Пишет "Голоса Всего 19 округ" и "Процент 19 Округ" в строку отчёта.

    В обычных строках нулевые голоса не выводятся (как и остальные колонки);
    в итоговых строках выводятся всегда, жирным шрифтом.
    """
    center = Alignment(horizontal="center")
    if votes > 0 or is_total_row:
        cell = sheet.cell(row, column, votes)
        cell.alignment = center
        if is_total_row:
            cell.font = Font(bold=True)
    percent = _okrug19_percent(votes, total)
    if percent:
        cell = sheet.cell(row, column + 1, percent)
        cell.alignment = center
        if is_total_row:
            cell.font = Font(bold=True)


def _percent_denominator_groups(
    short: bool = False, with_total: bool = True
) -> list[Optional[str]]:
    """
    Возвращает список, параллельный списку предикатов из _custom_groups.

    Описание:
        Для каждой колонки отчёта указывает, как считать процент в итоговых
        строках:
          - имя группы (например, "ДЭГ") — значит, процент для этой колонки
            нужно считать не от общего количества людей, а от количества
            людей, у которых сработал предикат "Планирует" этой же группы
            (см. GROUP_PLAN_PREDICATE). Так помечаются колонки
            "Проголосовал (QR-код)" и "Проголосовал (ответственный)".
          - None — процент считается как раньше, от общего количества людей
            (все остальные колонки, включая саму "Планирует" и "Всего").

    Аргументы:
        short: сокращённый набор колонок (см. _get_custom_columns).
        with_total: добавлять ли колонку "Всего" в сокращённом режиме.

    Возвращает:
        list: список той же длины и порядка, что и предикаты из
            _custom_groups(short, with_total).
    """
    bases: list[Optional[str]] = []
    for group, subs in _custom_groups(short, with_total):
        for sub, _predicate in subs:
            if (
                sub in ("Проголосовал (QR-код)", "Проголосовал (ответственный)")
                and group in GROUP_PLAN_PREDICATE
            ):
                bases.append(group)
            else:
                bases.append(None)
    return bases


def _get_custom_columns(
    short: bool = False, with_total: bool = True
) -> list[tuple[str, str, Callable[[Employee], bool]]]:
    """
    Возвращает список колонок отчёта — полный или сокращённый.

    Описание:
        При short=True в каждой группе способов голосования остаётся только
        по две подколонки:
          - ДЭГ: "Планирует" и "Проголосовал (QR-код)";
          - На участке: "Планирует" и "Проголосовал (QR-код)";
          - На участке УВЗ: "Заявление оформил" и "Проголосовал (QR-код)";
          - УИК-19: "Планирует" и "Проголосовал (QR-код)".
        Одиночные колонки "Не определился" и "Уважительная причина" убираются.

        Если with_total=True (в любом режиме, коротком или полном), последней
        добавляется колонка "Всего" — предикат p.voted. В отчёте "по людям"
        (см. custom_report) она не нужна и явно отключается (with_total=False).
        В отчёте "по производствам" (см. custom_production_summary) она нужна
        всегда: при разбивке по цехам колонка "Всего" считается по каждому
        цеху отдельно, а строка "Итого"/"Всего по Обществу" просто суммирует
        цеха — как и любая другая колонка-предикат, без специальной логики.
        Поскольку строка ИТОГО считает проценты через
        _format_with_percent(значение, total_people), эта колонка
        автоматически покажет процент проголосовавших от общего числа
        сотрудников (или от числа людей в цехе/производстве — в зависимости
        от отчёта).

    Аргументы:
        short: включает сокращённый режим колонок (см. описание выше).
        with_total: добавлять ли колонку "Всего" последней колонкой — и в
            коротком, и в полном режиме.

    Возвращает:
        list: список кортежей (имя_группы, имя_подколонки, предикат).
    """
    if not short:
        columns = list(CUSTOM_COLUMNS)
        if with_total:
            columns.append(("", "Всего", lambda p: p.voted))
        return columns

    allowed_subs_by_group = {
        "ДЭГ": {"Планирует", "Проголосовал (QR-код)"},
        "На участке": {"Планирует", "Проголосовал (QR-код)"},
        "На участке УВЗ": {"Заявление оформил", "Проголосовал (QR-код)"},
        "УИК-19": {"Планирует", "Проголосовал (QR-код)"},
    }
    columns = [
        (group, sub, predicate)
        for group, sub, predicate in CUSTOM_COLUMNS
        if sub in allowed_subs_by_group.get(group, set())
    ]
    if with_total:
        columns.append(("", "Всего", lambda p: p.voted))
    return columns


# ==============================================================================
# Фильтрация и заголовки
# ==============================================================================


def _custom_qs(params: dict) -> QuerySet:
    """
    Строит QuerySet сотрудников на основе параметров формы конструктора.

    Описание:
        Применяет фильтры, выбранные пользователем (производство, цех, способ,
        округ, УИК и т.д.), к общему списку сотрудников. Поддерживает специальные
        значения фильтров, такие как "Пусто" (none) и "20+21".

    Аргументы:
        params: словарь параметров из запроса (production, service, dep, method,
            where, okrug, uik).

    Возвращает:
        QuerySet: отфильтрованный список сотрудников.
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

    # Фильтр по запланированному способу голосования.
    if method == "none":
        qs = qs.filter(method="")
    elif method:
        qs = qs.filter(method=method)

    # Фильтр по фактическому месту голосования.
    if where == "none":
        qs = qs.filter(voted_method="")
    elif where:
        qs = qs.filter(voted_method=where)

    # Фильтр по избирательному округу.
    if okrug == "none":
        qs = qs.filter(okrug="")
    elif okrug == "20+21":
        qs = qs.filter(okrug__in=["20", "21"])
    elif okrug:
        qs = qs.filter(okrug=okrug)

    if uik:
        qs = qs.filter(uik=uik)

    return qs


def _custom_groups(
    short: bool = False, with_total: bool = True
) -> list[tuple[str, list[tuple[str, Callable]]]]:
    """
    Группирует колонки из CUSTOM_COLUMNS (или сокращённого набора) по имени группы.

    Описание:
        Преобразует плоский список колонок в иерархическую структуру для отрисовки
        объединённых заголовков в Excel. Колонки с одинаковым именем группы
        (идущие подряд) объединяются в одну группу.

    Колонки с пустым именем группы (одиночные) не сливаются между собой,
    а остаются независимыми столбцами. Это позволяет корректно
    отображать несколько одиночных колонок подряд (например, "Не определился"
    и "Уважительная причина").

    Аргументы:
        short: если True, используется сокращённый набор колонок из
            _get_custom_columns (см. её описание).
        with_total: добавлять ли колонку "Всего" в сокращённом режиме
            (передаётся в _get_custom_columns).

    Возвращает:
        list: список кортежей (имя_группы, [(имя_подколонки, предикат), ...]).
    """
    groups = []
    for group, sub, predicate in _get_custom_columns(short, with_total):
        # Пустая группа означает одиночную колонку без общего заголовка.
        # Такие колонки не должны объединяться с другими пустыми группами.
        if not group:
            groups.append((group, [(sub, predicate)]))
        # Если текущая колонка принадлежит той же группе, что и предыдущая,
        # добавляем её в конец списка подколонок этой группы.
        elif groups and groups[-1][0] == group:
            groups[-1][1].append((sub, predicate))
        else:
            # Иначе создаём новую группу.
            groups.append((group, [(sub, predicate)]))
    return groups


def _draw_custom_groups(
    sheet: Any,
    start_row: int,
    start_col: int,
    short: bool = False,
    with_total: bool = True,
) -> list[Callable]:
    """
    Отрисовывает объединённые заголовки групп и подколонки в листе Excel.

    Описание:
        Рисует шапку отчёта. Группы колонок объединяются по горизонтали в верхней строке.
        Подколонки рисуются во второй строке. Если имя группы пустое (одиночная колонка),
        ячейка объединяется по вертикали на две строки.

    Аргументы:
        sheet: объект листа openpyxl.
        start_row: номер строки, с которой начинается шапка.
        start_col: номер колонки, с которой начинается шапка.
        short: если True, используется сокращённый набор колонок
            (см. _get_custom_columns).
        with_total: добавлять ли колонку "Всего" в сокращённом режиме.

    Возвращает:
        list: список предикатов в порядке их следования в колонках листа.
    """
    bold = Font(bold=True)
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    column = start_col
    predicates = []

    for group, subs in _custom_groups(short, with_total):
        start, end = column, column + len(subs) - 1

        if group:
            # Объединяем ячейки для имени группы по горизонтали.
            sheet.merge_cells(
                start_row=start_row,
                start_column=start,
                end_row=start_row,
                end_column=end,
            )
            head = sheet.cell(start_row, column, group)
        else:
            # Для одиночной колонки (пустая группа) объединяем по вертикали.
            sheet.merge_cells(
                start_row=start_row,
                start_column=start,
                end_row=start_row + 1,
                end_column=end,
            )
            head = sheet.cell(start_row, column, subs[0][0])

        head.font = bold
        head.alignment = centered

        # Отрисовка подколонок и сбор предикатов.
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


# Ведущие колонки отчёта «по людям».
PEOPLE_LEAD_HEADERS = ("Номер УИК", "Цех", "Таб.№", "ФИО", "Округ")


def _people_table(
    sheet: Any, people: QuerySet, short: bool, with_total: bool
) -> tuple[int, list[int], dict[str, int], list[Optional[str]]]:
    """
    Шапка и строки отчёта «по людям»: сотрудник — строка, «+» в колонках,
    где выполняется условие колонки.

    Возвращает:
        tuple: (номер строки для итога, количество «+» по колонкам,
                число «Планирует» по группам, знаменатели процентов
                — см. _percent_denominator_groups).
    """
    for column, title in enumerate(PEOPLE_LEAD_HEADERS, 1):
        sheet.merge_cells(
            start_row=1, start_column=column, end_row=2, end_column=column
        )
        cell = sheet.cell(1, column, title)
        cell.font = BOLD
        cell.alignment = CENTERED
    first_mark_col = len(PEOPLE_LEAD_HEADERS) + 1
    predicates = _draw_custom_groups(
        sheet, start_row=1, start_col=first_mark_col, short=short, with_total=with_total
    )

    widths = (10, 8, 10, 42, 8) + (12,) * len(predicates)
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A3"

    row = 3
    totals = [0] * len(predicates)
    plan_counts: dict[str, int] = defaultdict(int)
    # iterator() с пачками по 2000 строк не держит всю выборку в памяти.
    for person in people.iterator(chunk_size=2000):
        for column, value in enumerate(
            (
                person.uik,
                person.department,
                person.tab_number,
                person.fio,
                person.okrug,
            ),
            1,
        ):
            sheet.cell(row, column, value)
        for offset, predicate in enumerate(predicates):
            if predicate(person):
                sheet.cell(row, first_mark_col + offset, "+")
                totals[offset] += 1
        for group, plan_predicate in GROUP_PLAN_PREDICATE.items():
            if plan_predicate(person):
                plan_counts[group] += 1
        row += 1

    return row, totals, plan_counts, _percent_denominator_groups(short, with_total)


def _people_totals(
    sheet: Any,
    row: int,
    people_count: int,
    totals: list[int],
    plan_counts: dict[str, int],
    denominators: list[Optional[str]],
) -> None:
    """Итоги по колонкам с «+»: число и процент (жирным, по центру)."""
    first_mark_col = len(PEOPLE_LEAD_HEADERS) + 1
    for offset, total in enumerate(totals):
        denom_group = denominators[offset]
        denom = plan_counts[denom_group] if denom_group else people_count
        cell = sheet.cell(
            row, first_mark_col + offset, _format_with_percent(total, denom)
        )
        cell.font = BOLD
        cell.alignment = CENTERED


def custom_report(params: dict, short: bool = False) -> Workbook:
    """
    Сводный Excel-отчёт «По людям» по фильтрам конструктора.

    Описание:
        Каждый сотрудник — отдельная строка, в колонках «+» по CUSTOM_COLUMNS.
        Внизу строка ИТОГО с количеством и процентами.

    Аргументы:
        params: параметры фильтров конструктора.
        short: если True, для каждого способа голосования остаются только
            две подколонки (см. _get_custom_columns), а «Не определился»
            и «Уважительная причина» убираются. Колонка «Всего» в этом
            отчёте не нужна.

    Возвращает:
        Workbook: книга openpyxl с готовым отчётом.
    """
    people = _custom_qs(params).order_by(
        "department", "surname", "name", "patronymic", "uik"
    )
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводный отчёт"

    row, totals, plans, denominators = _people_table(
        sheet, people, short, with_total=False
    )
    people_count = row - 3
    sheet.cell(row, 1, "ИТОГО").font = BOLD
    cell = sheet.cell(row, 2, people_count)
    cell.font = BOLD
    cell.alignment = CENTERED
    _people_totals(sheet, row, people_count, totals, plans, denominators)

    _apply_border(sheet)
    return book


def department_custom_report(department: str, params: dict) -> Workbook:
    """
    Отчёт «по людям» по одному цеху для архива конструктора (полный набор колонок).

    Аргументы:
        department: номер цеха.
        params: фильтры конструктора (производство, способ, округ и т.д.).

    Возвращает:
        Workbook: книга openpyxl с отчётом по цеху.
    """
    people = (
        _custom_qs(params)
        .filter(department=department)
        .order_by("surname", "name", "patronymic", "uik")
    )
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = f"Цех {department}"[:31]

    row, totals, plans, denominators = _people_table(
        sheet, people, False, with_total=True
    )
    people_count = row - 3
    # Подпись ИТОГО — под колонкой «ФИО», а не под «Номер УИК»/«Цех».
    cell = sheet.cell(row, 4, f"ИТОГО: {people_count} чел.")
    cell.font = BOLD
    cell.alignment = CENTERED
    _people_totals(sheet, row, people_count, totals, plans, denominators)

    _apply_border(sheet)
    return book


@dataclass
class _GroupStats:
    """
    Показатели группы людей: цеха, производства или всего общества.

    values — число людей по каждой колонке-предикату, plans — число
    «Планирует» по группам способов (знаменатель процентов «Проголосовал»),
    total19 и votes19 — колонки 19 округа.
    """

    people: int
    values: list[int]
    plans: dict[str, int]
    total19: int
    votes19: int

    @classmethod
    def empty(cls, columns: int) -> "_GroupStats":
        """Нулевые показатели — начальное значение для сумм."""
        return cls(0, [0] * columns, dict.fromkeys(GROUP_PLAN_PREDICATE, 0), 0, 0)

    @classmethod
    def of(cls, persons: list, predicates: list[Callable]) -> "_GroupStats":
        """Считает показатели по списку людей."""
        return cls(
            people=len(persons),
            values=[sum(1 for p in persons if pred(p)) for pred in predicates],
            plans={
                group: sum(1 for p in persons if plan(p))
                for group, plan in GROUP_PLAN_PREDICATE.items()
            },
            total19=sum(1 for p in persons if _okrug19_total(p)),
            votes19=sum(1 for p in persons if _okrug19_votes(p)),
        )

    def add(self, other: "_GroupStats") -> None:
        """Прибавляет показатели другой группы (для строк «Итого» и «Всего»)."""
        self.people += other.people
        self.values = [a + b for a, b in zip(self.values, other.values)]
        for group, count in other.plans.items():
            self.plans[group] += count
        self.total19 += other.total19
        self.votes19 += other.votes19


@dataclass
class _SummaryLayout:
    """Номера колонок и настройки отчёта «по производствам» конструктора."""

    count_col: int  # «Всего» людей
    total19_col: int  # «Всего 19 округ»
    pred_start: int  # первая колонка-предикат
    tail_start: int  # «Голоса Всего 19 округ»
    short: bool
    column_subs: list[str]  # имена подколонок (чтобы найти колонку «Всего»)
    denominators: list[Optional[str]]  # см. _percent_denominator_groups


def _write_summary_row(
    sheet: Any, row: int, layout: _SummaryLayout, stats: _GroupStats
) -> None:
    """
    Строка цеха или производства: числа без жирного шрифта, нули не выводятся.

    В полном режиме колонки «Проголосовал» показывают процент от «Планирует»
    этой же строки, колонка «Всего» — процент от числа людей в строке.
    """
    sheet.cell(row, layout.count_col, stats.people).alignment = CENTER
    sheet.cell(row, layout.total19_col, stats.total19).alignment = CENTER
    for offset, val in enumerate(stats.values):
        if val <= 0:
            continue
        denom_group = layout.denominators[offset]
        if not layout.short and denom_group:
            value = _format_with_percent(val, stats.plans[denom_group])
        elif not layout.short and layout.column_subs[offset] == "Всего":
            value = _format_with_percent(val, stats.people)
        else:
            value = val
        sheet.cell(row, layout.pred_start + offset, value).alignment = CENTER
    _write_okrug19_tail(sheet, row, layout.tail_start, stats.total19, stats.votes19)


def _write_summary_total(
    sheet: Any, row: int, label: str, layout: _SummaryLayout, stats: _GroupStats
) -> None:
    """Строка «Итого» или «Всего по Обществу»: жирным, во всех колонках — процент."""
    sheet.cell(row, layout.count_col - 1, label).font = BOLD
    for col, value in (
        (layout.count_col, stats.people),
        (layout.total19_col, stats.total19),
    ):
        cell = sheet.cell(row, col, value)
        cell.font = BOLD
        cell.alignment = CENTER
    for offset, val in enumerate(stats.values):
        denom_group = layout.denominators[offset]
        denom = stats.plans[denom_group] if denom_group else stats.people
        cell = sheet.cell(
            row, layout.pred_start + offset, _format_with_percent(val, denom)
        )
        cell.font = BOLD
        cell.alignment = CENTERED
    _write_okrug19_tail(
        sheet, row, layout.tail_start, stats.total19, stats.votes19, is_total_row=True
    )


def custom_production_summary(
    params: dict,
    include_depts: bool = False,
    moment: Optional[datetime] = None,
    short: bool = False,
) -> Workbook:
    """
    Генерирует сводный отчёт с группировкой по производствам.

    Описание:
        Агрегирует данные по производствам (опционально с разбивкой по цехам).
        Показывает количество сотрудников по каждому производству и процентные
        соотношения по колонкам из конструктора. Производство — поле service
        (см. reports/production.py).

    Аргументы:
        params: параметры фильтров конструктора.
        include_depts: если True, выводит строки по цехам внутри каждого
            производства и строку «Итого» по производству.
        moment: момент времени для заголовка отчёта.
        short: если True, для каждого способа голосования оставляются только
            подколонки "Планирует" и "Проголосовал (QR-код)", а колонки
            "Не определился" и "Уважительная причина" заменяются одной
            колонкой "Всего" — в строках "Итого"/"Всего по Обществу" она
            покажет процент проголосовавших от общего числа сотрудников.
            Если False (полный режим), колонки "Проголосовал" в КАЖДОЙ строке
            показывают процент от числа людей, планирующих голосовать этим
            же способом в этой строке (см. GROUP_PLAN_PREDICATE), а колонка
            "Всего" — процент от числа людей в строке.

        Колонки 19 округа: "Всего 19 округ" — после "Всего"; "Голоса Всего
        19 округ" и "Процент 19 Округ" — в конце таблицы (см. _okrug19_total
        и _okrug19_votes).

    Возвращает:
        Workbook: книга openpyxl со сводным отчётом.
    """
    moment = moment or timezone.localtime()
    people = _custom_qs(params).order_by(
        "service", "department", "surname", "name", "patronymic"
    )

    # Плоский список предикатов и имён подколонок (нужны, чтобы найти «Всего»).
    predicates, column_subs = [], []
    for _group, subs in _custom_groups(short):
        for sub, predicate in subs:
            predicates.append(predicate)
            column_subs.append(sub)

    leads = lead_headers(include_depts)
    pred_start = len(leads) + 1
    layout = _SummaryLayout(
        count_col=len(leads) - 1,
        total19_col=len(leads),
        pred_start=pred_start,
        tail_start=pred_start + len(predicates),
        short=short,
        column_subs=column_subs,
        denominators=_percent_denominator_groups(short, with_total=True),
    )
    total_columns = len(leads) + len(predicates) + len(OKRUG19_TAIL_HEADERS)

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сводный отчёт"

    # Шапка: заголовок, ведущие колонки, группы способов, колонки 19 округа.
    write_title(sheet, moment, total_columns)
    column = draw_lead_headers(sheet, leads)
    _draw_custom_groups(sheet, start_row=HEADER_ROW, start_col=column, short=short)
    _draw_okrug19_tail_headers(sheet, HEADER_ROW, layout.tail_start)

    lead_widths = (16, 14, 10, 12) if include_depts else (24, 10, 12)
    widths = lead_widths + (12,) * len(predicates) + (14, 12)
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A4"

    data = group_people(people.iterator(chunk_size=2000), include_depts)
    grand = _GroupStats.empty(len(predicates))
    row = FIRST_DATA_ROW

    for production in production_order(data):
        if not include_depts:
            # Одна строка на производство.
            stats = _GroupStats.of(data[production], predicates)
            sheet.cell(row, 1, production)
            _write_summary_row(sheet, row, layout, stats)
            grand.add(stats)
            row += 1
            continue

        # Название производства, строки его цехов и строка «Итого».
        write_production_heading(sheet, row, production, total_columns)
        row += 1
        subtotal = _GroupStats.empty(len(predicates))
        for department in sorted(data[production], key=_by_number):
            stats = _GroupStats.of(data[production][department], predicates)
            sheet.cell(row, 1, production)
            sheet.cell(row, 2, padded_number(department))
            _write_summary_row(sheet, row, layout, stats)
            subtotal.add(stats)
            row += 1
        _write_summary_total(sheet, row, "Итого", layout, subtotal)
        grand.add(subtotal)
        row += 1

    _write_summary_total(sheet, row, "Всего по Обществу", layout, grand)
    _apply_border(sheet)
    return book
