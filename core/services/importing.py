"""
Модуль импорта и обновления данных сотрудников из Excel-файлов.

Описание:
    Содержит логику чтения выгрузок из Excel и записи данных в базу.
    Поддерживает сценарии импорта основной базы, отчётов штаба,
    массовую простановку явки и обработку сводных кастомных отчётов.
"""

import io
import zipfile
from collections.abc import Callable, Iterator
from datetime import date, datetime, time
from typing import Any

from django.db.models import F
from django.utils import timezone

from ..models import DEG, UIK, UIK19, UVZ, Employee
from .excel import BATCH, COLUMNS, _date, _header, _sheet, _text

# Сообщение для загрузки, которая не является ZIP-архивом.
BAD_ZIP = "Ошибка: файл не является корректным ZIP-архивом"

# ==============================================================================
# Импорт основной базы
# ==============================================================================


def _rows_by_tab(rows: Iterator, positions: dict) -> dict:
    """
    Парсит строки Excel-файла в словарь, ключом которого является табельный номер.

    Описание:
        Преобразует типы данных в соответствии с картой колонок COLUMNS.
        Табельные номера, состоящие только из цифр, дополняются нулями до 7 знаков.

    Аргументы:
        rows: итератор строк Excel-файла.
        positions: словарь соответствия имен колонок их индексам.

    Возвращает:
        dict: словарь вида {tab_number: {field: value, ...}}.
    """
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
        if tab and str(tab).isdigit():
            tab = str(tab).zfill(7)

        if tab:
            parsed[tab] = values

    return parsed


def _known_rows(tabs: list, fields: list) -> dict:
    """
    Получает из базы данных существующие записи сотрудников по их табельным номерам.

    Описание:
        Использует чанки по 2000 записей для оптимизации нагрузки на базу данных
        и предотвращения превышения лимита параметров SQL-запроса.

    Аргументы:
        tabs: список табельных номеров для поиска.
        fields: список полей модели, которые необходимо извлечь.

    Возвращает:
        dict: словарь вида {tab_number: {pk: id, field: value, ...}}.
    """
    known = {}
    tabs_list = list(tabs)

    for start in range(0, len(tabs_list), 2000):
        chunk = tabs_list[start : start + 2000]
        for row in Employee.objects.filter(tab_number__in=chunk).values(
            "pk", "tab_number", *fields
        ):
            known[row.pop("tab_number")] = row

    return known


def import_base(upload: Any) -> tuple[int, int, int]:
    """
    Импортирует или обновляет основную базу сотрудников из Excel-файла.

    Возвращает:
        tuple[int, int, int]: (количество созданных, количество обновленных, всего обработано).
    """
    with _sheet(upload) as rows:
        positions = _header(rows, COLUMNS)
        if "Таб№" not in positions:
            raise ValueError("Ошибка: не найдена колонка 'Таб№'")
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


# ==============================================================================
# Способы голосования из отчёта штаба
# ==============================================================================

# Заголовок колонки (часть текста) -> код способа; в выбранной колонке стоит «1»
CHOICE_COLUMNS = {
    "Очно на своем избирательном участке": UIK,
    "При помощи дистанционного электронного голосования ДЭГ": DEG,
    "Очно на временном избирательном участке на территории предприятия": UVZ,
    "Очно на избирательном участке 19 округа": UIK19,
}


def _choice_header(rows: list) -> tuple[int, int, dict[str, int]]:
    """Строка шапки: (номер строки, колонка «Таб.№», {способ: колонка})."""
    for index, row in enumerate(rows):
        cells = [_text(cell) for cell in row]
        if "Таб.№" not in cells and "ФИО" not in cells:
            continue
        if "Таб.№" not in cells:
            raise ValueError("Ошибка: не найдена обязательная колонка 'Таб.№'")
        methods = {
            method: column
            for column, cell in enumerate(cells)
            for title, method in CHOICE_COLUMNS.items()
            if title in cell
        }
        if not methods:
            raise ValueError("Ошибка: не найдены колонки способов голосования")
        return index, cells.index("Таб.№"), methods
    raise ValueError("Ошибка: не найдена строка с заголовками")


def import_voting_choices(upload: Any) -> tuple[int, int, int]:
    """
    Импортирует выбранные способы голосования из отчёта штаба (поле method).

    В строке должна стоять «1» ровно в одной колонке способа; две и больше
    отметок — ошибка, ни одной — строка пропускается.

    Возвращает:
        tuple[int, int, int]: (обновлено, всего строк, ошибок/пропусков).
    """
    with _sheet(upload) as rows:
        all_rows = list(rows)
    if not all_rows:
        raise ValueError("Ошибка: файл пустой")

    header_index, tab_column, method_columns = _choice_header(all_rows)
    chosen, total, errors = {}, 0, 0
    for row in all_rows[header_index + 1 :]:
        if not any(row):
            continue
        total += 1
        tab = _tab(row[tab_column]) if tab_column < len(row) else ""
        marked = [
            method
            for method, column in method_columns.items()
            if column < len(row) and _text(row[column]) == "1"
        ]
        if len(marked) > 1:
            errors += 1
        elif tab and marked:
            chosen[tab] = marked[0]

    changed = [
        employee
        for employee in Employee.objects.filter(tab_number__in=chosen)
        if employee.method != chosen[employee.tab_number]
    ]
    for employee in changed:
        employee.method = chosen[employee.tab_number]
    Employee.objects.bulk_update(changed, fields=["method"], batch_size=BATCH)

    found = Employee.objects.filter(tab_number__in=chosen).count()
    return len(changed), total, errors + len(chosen) - found


# ==============================================================================
# Отметки явки (базовые)
# ==============================================================================


def set_turnout(queryset: Any, voted: bool = True) -> int:
    """
    Массовое обновление статуса явки для переданного QuerySet.

    Описание:
        При отметке явки также фиксируется текущее время и копируется
        запланированный способ голосования в поле voted_method.

    Аргументы:
        queryset: QuerySet сотрудников для обновления.
        voted: булево значение явки (по умолчанию True).

    Возвращает:
        int: количество обновленных записей.
    """
    return queryset.update(
        voted=voted,
        voted_at=timezone.now() if voted else None,
        voted_method=F("method") if voted else "",
    )


def mark_voted(tabs: list, voted: bool = True) -> tuple[int, int]:
    """
    Отмечает явку по списку табельных номеров.

    Аргументы:
        tabs: список табельных номеров.
        voted: булево значение явки (по умолчанию True).

    Возвращает:
        tuple[int, int]: (количество обновленных записей, количество не найденных в базе).
    """
    tabs_set = {str(t) for t in tabs if t}
    if not tabs_set:
        return 0, 0

    formatted_tabs = {t.zfill(7) if t.isdigit() else t for t in tabs_set}

    found = Employee.objects.filter(tab_number__in=formatted_tabs)
    missing = len(formatted_tabs) - found.count()

    return set_turnout(found, voted), missing


# ==============================================================================
# Импорт явки из отчёта штаба (с датой и временем)
# ==============================================================================


def _combine_datetime(date_val: Any, time_val: Any) -> datetime:
    """
    Объединяет дату и время из ячеек Excel в timezone-aware datetime объект.

    Описание:
        Поддерживает как нативные типы данных openpyxl, так и строковые представления.
    """
    if not date_val and not time_val:
        return timezone.now()

    if isinstance(date_val, datetime) and isinstance(time_val, time):
        dt = datetime.combine(date_val.date(), time_val)
    elif isinstance(date_val, date) and isinstance(time_val, time):
        dt = datetime.combine(date_val, time_val)
    else:
        d_str = str(date_val).strip() if date_val else ""
        t_str = str(time_val).strip() if time_val else ""
        combined_str = f"{d_str} {t_str}".strip()

        formats = [
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%d.%m.%Y %H:%M:%S",
            "%d.%m.%Y %H:%M",
            "%d.%m.%y %H:%M:%S",
            "%d.%m.%y %H:%M",
        ]
        dt = None
        for fmt in formats:
            try:
                dt = datetime.strptime(combined_str, fmt)
                break
            except ValueError:
                continue

        if not dt:
            return timezone.now()

    if timezone.is_naive(dt):
        dt = timezone.make_aware(dt)
    return dt


def _workbooks(upload: Any) -> Iterator[Any]:
    """
    Книги Excel из загруженного файла: сам файл (.xlsx) или все .xlsx из ZIP.

    Служебные папки macOS (__MACOSX) и файлы других типов в архиве пропускаются.

    Исключения:
        ValueError: если файл с расширением .zip не является ZIP-архивом.
    """
    if not upload.name.lower().endswith(".zip"):
        yield upload
        return
    try:
        archive = zipfile.ZipFile(upload, "r")
    except zipfile.BadZipFile as exc:
        raise ValueError(BAD_ZIP) from exc
    with archive:
        for info in archive.infolist():
            name = info.filename
            if name.endswith(".xlsx") and not name.startswith("__MACOSX"):
                yield io.BytesIO(archive.read(name))


def _tab(value: Any) -> str:
    """Табельный номер из ячейки: только цифры, дополненные нулями до 7 знаков; иначе ""."""
    tab = _text(value)
    return tab.zfill(7) if tab.isdigit() else ""


def import_turnout_hq_archive(upload: Any) -> tuple[int, int, int]:
    """
    Импорт отметок явки из ZIP-архива отчётов штаба с датой и временем.

    Описание:
        В каждой книге первая строка — шапка, дальше по колонкам:
        0 — табельный номер, 4 — дата, 5 — время голосования.
        Если у сотрудника не выбран способ голосования, при отметке явки
        ему ставится UIK.

    Возвращает:
        tuple[int, int, int]: (изменено, всего строк, ошибок/пропусков).
    """
    total_rows = total_errors = 0
    voted_at = {}  # табельный номер -> время голосования

    for workbook in _workbooks(upload):
        with _sheet(workbook) as rows:
            all_rows = list(rows)
        for row in all_rows[1:]:
            if not any(row):
                continue
            total_rows += 1
            tab = _tab(row[0])
            if not tab:
                total_errors += 1
                continue
            date_val = row[4] if len(row) > 4 else None
            time_val = row[5] if len(row) > 5 else None
            voted_at[tab] = _combine_datetime(date_val, time_val)

    if not voted_at:
        return 0, total_rows, total_errors

    found = list(Employee.objects.filter(tab_number__in=voted_at))
    # Ошибки — табельные номера, которых нет в базе. Люди без способа голосования
    # ошибкой не считаются: им ставится УИК и они отмечаются как проголосовавшие
    total_errors += len(voted_at) - len(found)

    for emp in found:
        emp.voted = True
        emp.voted_at = voted_at[emp.tab_number]
        emp.method = emp.method or UIK
        emp.voted_method = emp.method
    if found:
        Employee.objects.bulk_update(
            found,
            fields=["voted", "voted_at", "voted_method", "method"],
            batch_size=BATCH,
        )
    return len(found), total_rows, total_errors


# ==============================================================================
# Импорт кастомного сводного отчёта
# ==============================================================================

# Колонки отчёта: (ключ, условие по названию группы и подколонки в нижнем регистре).
# Проверяются по порядку, колонке достаётся первый подошедший ключ.
CUSTOM_REPORT_COLUMNS: list[tuple[str, Callable[[str, str], bool]]] = [
    ("tab", lambda g, s: "таб.№" in g + s or "табельный" in g + s),
    ("plan_deg", lambda g, s: "дэг" in g and "планирует" in s),
    ("mark_deg", lambda g, s: "дэг" in g and "зарегистрирован" in s),
    ("voted_deg", lambda g, s: "дэг" in g and "проголосовал" in s),
    (
        "plan_uik",
        lambda g, s: "на участке" in g and "увз" not in g and "планирует" in s,
    ),
    (
        "voted_uik",
        lambda g, s: "на участке" in g and "увз" not in g and "проголосовал" in s,
    ),
    ("plan_uvz", lambda g, s: "увз" in g and "планирует" in s),
    ("voted_uvz", lambda g, s: "увз" in g and "проголосовал" in s),
    ("plan_u19", lambda g, s: "19" in g and "планирует" in s),
    ("detached", lambda g, s: "19" in g and "открепился" in s),
    ("voted_u19", lambda g, s: "19" in g and "проголосовал" in s),
    ("absence", lambda g, s: "уважительная" in g + s or "причина" in g + s),
    ("not_going", lambda g, s: "не пойдет" in g + s or "не пойдёт" in g + s),
]

# Колонки «Планирует» и «Проголосовал» -> способ голосования (первая отметка побеждает).
PLAN_COLUMNS = (
    ("plan_deg", DEG),
    ("plan_uik", UIK),
    ("plan_uvz", UVZ),
    ("plan_u19", UIK19),
)
VOTED_COLUMNS = (
    ("voted_deg", DEG),
    ("voted_uik", UIK),
    ("voted_uvz", UVZ),
    ("voted_u19", UIK19),
)

# Отметки, при которых «Уважительная причина» из файла не ставится.
MARKS_OVER_ABSENCE = ("method", "mark_deg", "detached", "not_going", "voted")


def _custom_report_columns(group_row: tuple, sub_row: tuple) -> dict[str, int]:
    """
    Номера колонок отчёта по двухстрочной шапке (группа + подколонка).

    Название группы стоит в объединённой ячейке только над первой подколонкой,
    поэтому оно протягивается вправо до следующего непустого названия.
    """
    columns = {}
    group = ""
    for index, cell in enumerate(group_row):
        group = str(cell or "").strip().lower() or group
        sub = (
            str((sub_row[index] if index < len(sub_row) else None) or "")
            .strip()
            .lower()
        )
        for key, matches in CUSTOM_REPORT_COLUMNS:
            if matches(group, sub):
                columns[key] = index
                break
    return columns


def _has_mark(row: tuple, index: int | None) -> bool:
    """Есть ли отметка в ячейке: любое непустое значение (так задумано)."""
    return (
        index is not None and index < len(row) and bool(str(row[index] or "").strip())
    )


def _row_changes(row: tuple, columns: dict[str, int]) -> dict[str, Any]:
    """Поля сотрудника, которые нужно изменить по одной строке отчёта."""
    marked = lambda key: _has_mark(row, columns.get(key))  # noqa: E731
    changes: dict[str, Any] = {}

    method = next((m for key, m in PLAN_COLUMNS if marked(key)), None)
    if method:
        changes["method"] = method
    if marked("mark_deg"):
        changes.update(method=DEG, mark_deg=True)
    if marked("detached"):
        changes.update(method=UIK19, detached=True)
    if marked("not_going"):
        changes["not_going"] = True
    if marked("absence"):
        changes["absence"] = True

    voted_method = next((m for key, m in VOTED_COLUMNS if marked(key)), None)
    if voted_method:
        changes.update(voted=True, voted_method=voted_method)
    return changes


def _save_changes(updates: dict[str, dict]) -> tuple[int, int]:
    """
    Записывает изменения в базу.

    Обновляются только реально изменившиеся сотрудники и только поля,
    которые встречаются в файле, — остальные данные не перезаписываются.

    Возвращает:
        tuple[int, int]: (изменено сотрудников, табельных номеров нет в базе).
    """
    existing = {
        e.tab_number: e for e in Employee.objects.filter(tab_number__in=updates)
    }
    to_update = []
    for tab, fields in updates.items():
        emp = existing.get(tab)
        if emp is None:
            continue
        changed = [f for f, value in fields.items() if getattr(emp, f) != value]
        for field in changed:
            setattr(emp, field, fields[field])
        if changed:
            to_update.append(emp)

    if to_update:
        fields_in_file = list(
            {field for fields in updates.values() for field in fields}
        )
        Employee.objects.bulk_update(to_update, fields=fields_in_file, batch_size=BATCH)
    return len(to_update), len(updates) - len(existing)


def _import_custom_sheet(all_rows: list) -> tuple[int, int, int]:
    """
    Импорт одного листа кастомного отчёта.

    Возвращает:
        tuple[int, int, int]: (изменено, строк данных, ошибок).
    """
    if len(all_rows) < 3:
        return 0, 0, 0

    # Шапка — первая строка с табельным номером, под ней строка подколонок.
    header = next(
        (
            i
            for i, row in enumerate(all_rows)
            if any(
                "таб.№" in str(c).lower() or "табельный" in str(c).lower()
                for c in row
                if c
            )
        ),
        0,
    )
    columns = _custom_report_columns(all_rows[header], all_rows[header + 1])
    if "tab" not in columns:
        return 0, 0, 1  # лист без табельных номеров

    updates: dict[str, dict] = {}
    rows = errors = 0
    for row in all_rows[header + 2 :]:
        if not any(row):
            continue
        rows += 1
        tab = _tab(row[columns["tab"]]) if columns["tab"] < len(row) else ""
        if not tab:
            errors += 1
            continue
        updates.setdefault(tab, {}).update(_row_changes(row, columns))

    for fields in updates.values():
        if any(fields.get(mark) for mark in MARKS_OVER_ABSENCE):
            fields.pop("absence", None)

    changed, missing = _save_changes(updates) if updates else (0, 0)
    return changed, rows, errors + missing


def import_custom_report_archive(upload: Any) -> tuple[int, int, int]:
    """
    Импортирует данные из кастомного сводного отчёта (Excel или ZIP с Excel).

    Описание:
        Распознаёт двухстрочную шапку с объединёнными ячейками (группа +
        подколонка) и обновляет поля сотрудника: способ, явку, служебные
        отметки. Каждая книга обрабатывается и сохраняется отдельно.

    Возвращает:
        tuple[int, int, int]: (изменено, всего строк, ошибок/пропусков).
    """
    totals = [0, 0, 0]
    for workbook in _workbooks(upload):
        with _sheet(workbook) as rows:
            result = _import_custom_sheet(list(rows))
        totals = [a + b for a, b in zip(totals, result, strict=True)]
    return tuple(totals)


# ==============================================================================
# Импорт отметок "Голосование через ответственного"
# ==============================================================================

RESPONSIBLE_COLUMNS = {
    "Цех": "department",
    "Таб": "tab_number",
    "ФИО": "fio",
    "Отметка": "mark",
}

# Другие написания заголовков, которые встречаются в файлах цехов
# (в том числе набранные в английской раскладке: «nf,tkmysq» = «табельный»).
HEADER_ALIASES = {
    "Таб": [
        "таб",
        "табельный",
        "табельный номер",
        "таб.",
        "таб №",
        "таб. №",
        "№ таб",
        "табельный №",
        "nf,tkmysq",
        "nf,",
        "nf,tkmysq yjvth",
    ],
    "Отметка": [
        "отметка",
        "ометка",
        "отм.",
        "проголосовал",
        "проголосовал (если проголосовал то 1)",
        "голосовал",
        "гол",
    ],
}


def _responsible_tabs(workbook: Any) -> tuple[set[str], int, int]:
    """
    Табельные номера с отметкой из одной книги.

    Возвращает:
        tuple: (номера с непустой отметкой, строк данных, ошибок).
    """
    with _sheet(workbook) as rows:
        try:
            positions = _header(rows, RESPONSIBLE_COLUMNS, aliases=HEADER_ALIASES)
        except ValueError:
            return set(), 0, 1
        tab_idx, mark_idx = positions.get("Таб"), positions.get("Отметка")
        if tab_idx is None or mark_idx is None:
            return set(), 0, 1

        marked, total, errors = set(), 0, 0
        for row in rows:
            if not any(row):
                continue
            total += 1
            tab = _text(row[tab_idx] if tab_idx < len(row) else None)
            if not tab:
                errors += 1
                continue
            # Отметка — любое значение, в том числе 0 (ячейка «0» тоже считается)
            if _text(row[mark_idx] if mark_idx < len(row) else None):
                marked.add(tab.zfill(7) if tab.isdigit() else tab)
    return marked, total, errors


def import_responsible_marks(upload: Any) -> tuple[int, int, int]:
    """
    Импортирует отметки «Голосование через ответственного» из Excel или ZIP.

    Описание:
        Ищет колонки «Таб» и «Отметка». Если в колонке «Отметка» есть любое
        непустое значение, сотруднику ставится proxy_vote. Время отметки
        записывается только тем, у кого его ещё нет (у проголосовавших
        по QR-коду время голосования не затирается).

    Возвращает:
        tuple[int, int, int]: (изменено, всего строк, ошибок/пропусков).
    """
    marked_tabs: set[str] = set()
    total_rows = total_errors = 0
    for workbook in _workbooks(upload):
        tabs, rows, errors = _responsible_tabs(workbook)
        marked_tabs |= tabs
        total_rows += rows
        total_errors += errors

    if not marked_tabs:
        return 0, total_rows, total_errors

    found = Employee.objects.filter(tab_number__in=marked_tabs)
    changed = found.update(proxy_vote=True)
    found.filter(voted_at__isnull=True).update(voted_at=timezone.now())
    total_errors += len(marked_tabs) - found.count()
    return changed, total_rows, total_errors
