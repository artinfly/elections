"""
Модуль импорта и обновления данных сотрудников из Excel-файлов.

Описание:
    Содержит логику чтения выгрузок из Excel и записи данных в базу.
    Поддерживает сценарии импорта базы, отчётов штаба и массовую простановку явки.
"""

import io
import zipfile
from datetime import date, datetime, time
from typing import Any, Iterator

from django.db.models import F
from django.utils import timezone

from .helpers import BATCH, COLUMNS, _date, _header, _sheet, _text
from .models import Employee

# ==============================================================================
# Импорт основной базы
# ==============================================================================


def _rows_by_tab(rows: Iterator, positions: dict) -> dict:
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
    known = {}
    tabs = list(tabs)
    for start in range(0, len(tabs), 2000):
        chunk = tabs[start : start + 2000]
        for row in Employee.objects.filter(tab_number__in=chunk).values(
            "pk", "tab_number", *fields
        ):
            known[row.pop("tab_number")] = row
    return known


def import_base(upload: Any) -> tuple[int, int, int]:
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
# Отметки явки (базовые)
# ==============================================================================


def set_turnout(queryset: Any, voted: bool = True) -> int:
    return queryset.update(
        voted=voted,
        voted_at=timezone.now() if voted else None,
        voted_method=F("method") if voted else "",
    )


def mark_voted(tabs: list, voted: bool = True) -> tuple[int, int]:
    tabs = {str(t) for t in tabs if t}
    if not tabs:
        return 0, 0

    # Приводим к формату с нулями, если это цифры
    formatted_tabs = {t.zfill(7) if t.isdigit() else t for t in tabs}

    found = Employee.objects.filter(tab_number__in=formatted_tabs)
    missing = len(formatted_tabs) - found.count()
    return set_turnout(found, voted), missing


# ==============================================================================
# Импорт отчётов штаба (Способы голосования)
# ==============================================================================

NEW_FORMAT_COLUMNS = {
    "Таб.№": "tab_number",
    "ФИО": "fio_raw",
    "Очно на своем избирательном участке": "uik",
    "При помощи дистанционного электронного голосования ДЭГ": "deg",
    "Очно на временном избирательном участке на территории предприятия": "uvz",
    "Очно на избирательном участке 19 округа": "u19",
}


def import_voting_choices(upload: Any) -> tuple[int, int, int]:
    with _sheet(upload) as rows:
        all_rows = list(rows)

    if not all_rows:
        raise ValueError("Ошибка: файл пустой")

    header_row_idx = 0
    for idx, row in enumerate(all_rows):
        cells = [str(cell or "").strip() for cell in row if cell is not None]
        if "Таб.№" in cells or "ФИО" in cells:
            header_row_idx = idx
            break

    if header_row_idx >= len(all_rows) - 1:
        raise ValueError("Ошибка: не найдена строка с данными")

    header_row = all_rows[header_row_idx]
    col_indices = {}
    for idx, cell in enumerate(header_row):
        cell_str = str(cell or "").strip()
        for key in NEW_FORMAT_COLUMNS:
            if key in cell_str:
                col_indices[key] = idx
                break

    if "Таб.№" not in col_indices or "ФИО" not in col_indices:
        raise ValueError("Ошибка: не найдены обязательные колонки 'Таб.№' или 'ФИО'")

    method_cols = {
        "uik": col_indices.get("Очно на своем избирательном участке"),
        "deg": col_indices.get(
            "При помощи дистанционного электронного голосования ДЭГ"
        ),
        "uvz": col_indices.get(
            "Очно на временном избирательном участке на территории предприятия"
        ),
        "u19": col_indices.get("Очно на избирательном участке 19 округа"),
    }

    if not any(method_cols.values()):
        raise ValueError("Ошибка: не найдены колонки способов голосования")

    file_data = {}
    total = 0
    parse_errors = 0

    for row in all_rows[header_row_idx + 1 :]:
        if not any(row):
            continue

        total += 1

        try:
            tab_col = col_indices["Таб.№"]
            tab_number = _text(row[tab_col] if tab_col < len(row) else None)
            if not tab_number:
                continue

            if tab_number.isdigit():
                tab_number = tab_number.zfill(7)

            selected_method = ""
            for method_code, col_idx in method_cols.items():
                if col_idx is not None and col_idx < len(row):
                    if _text(row[col_idx]) == "1":
                        if selected_method:
                            parse_errors += 1
                            selected_method = ""
                            break
                        selected_method = method_code

            if not selected_method:
                continue

            file_data[tab_number] = selected_method

        except Exception:
            parse_errors += 1
            continue

    if not file_data:
        return 0, total, parse_errors

    existing_employees = {
        emp.tab_number: emp
        for emp in Employee.objects.filter(tab_number__in=file_data.keys())
    }

    employees_to_update = []
    not_found_errors = 0

    for tab_number, selected_method in file_data.items():
        employee = existing_employees.get(tab_number)
        if employee is None:
            not_found_errors += 1
            continue

        if employee.method != selected_method:
            employee.method = selected_method
            employees_to_update.append(employee)

    if employees_to_update:
        Employee.objects.bulk_update(
            employees_to_update, fields=["method"], batch_size=BATCH
        )

    updated = len(employees_to_update)
    errors = parse_errors + not_found_errors

    return updated, total, errors


# ==============================================================================
# Импорт явки из списка табельных номеров (простой)
# ==============================================================================


def import_turnout(upload: Any) -> tuple[int, int, int]:
    with _sheet(upload) as rows:
        all_rows = list(rows)

    if not all_rows:
        raise ValueError("Ошибка: файл пустой")

    header = str(all_rows[0][0] or "").strip().lower()
    if "табель" not in header and "таб" not in header:
        raise ValueError("Ошибка: в А1 ожидается заголовок 'Табельный'")

    tabs = []
    for row in all_rows[1:]:
        if not row:
            continue
        tab = _text(row[0])
        if tab and str(tab).isdigit():
            tabs.append(str(tab).zfill(7))

    if not tabs:
        return 0, 0, 0

    total_rows = len(tabs)
    found = Employee.objects.filter(tab_number__in=tabs)
    missing_in_db = total_rows - found.count()

    valid_found = found.exclude(method="")
    skipped_no_method = found.filter(method="").count()

    changed = set_turnout(valid_found, voted=True)
    errors = missing_in_db + skipped_no_method

    return changed, total_rows, errors


# ==============================================================================
# Импорт явки из отчёта штаба (с датой и временем)
# ==============================================================================


def _combine_datetime(date_val: Any, time_val: Any) -> datetime:
    """Объединяет дату и время из ячеек Excel в timezone-aware datetime."""
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


def import_turnout_hq_archive(upload: Any) -> tuple[int, int, int]:
    """
    Импорт отметок явки из файла штаба с датой и временем.
    1 колонка (индекс 0): табельный номер
    4 колонка (индекс 4): дата
    5 колонка (индекс 5): время
    """
    total_changed = 0
    total_rows = 0
    total_errors = 0
    all_tabs_data = {}

    try:
        with zipfile.ZipFile(upload, "r") as zip_ref:
            for file_info in zip_ref.infolist():
                if file_info.filename.endswith(
                    ".xlsx"
                ) and not file_info.filename.startswith("__MACOSX"):
                    with zip_ref.open(file_info) as excel_file:
                        file_content = io.BytesIO(zip_ref.read(file_info.filename))
                        with _sheet(file_content) as rows:
                            all_rows = list(rows)

                        if len(all_rows) < 2:
                            continue

                        for row in all_rows[1:]:
                            if not any(row):
                                continue
                            total_rows += 1
                            try:
                                tab = _text(row[0] if len(row) > 0 else None)
                                if not tab or not str(tab).isdigit():
                                    total_errors += 1
                                    continue
                                tab = str(tab).zfill(7)

                                date_val = row[4] if len(row) > 4 else None
                                time_val = row[4] if len(row) > 4 else None

                                all_tabs_data[tab] = _combine_datetime(
                                    date_val, time_val
                                )

                            except Exception:
                                total_errors += 1
    except zipfile.BadZipFile:
        raise ValueError("Ошибка: файл не является корректным ZIP-архивом.")

    if not all_tabs_data:
        return 0, total_rows, total_errors

    tabs_list = list(all_tabs_data.keys())
    found = Employee.objects.filter(tab_number__in=tabs_list)

    valid_found = found.exclude(method="")
    total_errors += (len(tabs_list) - found.count()) + found.filter(method="").count()

    employees_to_update = []
    for emp in valid_found:
        emp.voted = True
        emp.voted_at = all_tabs_data[emp.tab_number]
        emp.voted_method = emp.method
        employees_to_update.append(emp)

    if employees_to_update:
        Employee.objects.bulk_update(
            employees_to_update,
            fields=["voted", "voted_at", "voted_method"],
            batch_size=BATCH,
        )

    return len(employees_to_update), total_rows, total_errors


def import_custom_report_archive(upload: Any) -> tuple[int, int, int]:
    total_changed = 0
    total_rows = 0
    total_errors = 0

    def _get_col_indices(header_rows: list) -> dict:
        indices = {}
        for c_idx in range(len(header_rows[0])):
            group = str(header_rows[0][c_idx] or "").strip().lower()
            sub = str(header_rows[1][c_idx] or "").strip().lower()
            combined = f"{group}{sub}"

            if "таб.№" in combined or "табельный" in combined:
                indices["tab"] = c_idx
            elif "дэг" in group and "планирует" in sub:
                indices["plan_deg"] = c_idx
            elif "дэг" in group and "зарегестрирован" in sub:
                indices["mark_deg"] = c_idx
            elif "дэг" in group and "проголосовал" in sub:
                indices["voted_deg"] = c_idx
            elif "на участке" in group and "увз" not in group and "планирует" in sub:
                indices["plan_uik"] = c_idx
            elif "на участке" in group and "увз" not in group and "проголосовал" in sub:
                indices["voted_uik"] = c_idx
            elif "увз" in group and "планирует" in sub:
                indices["plan_uvz"] = c_idx
            elif "увз" in group and "заявление" in sub:
                indices["mark_uvz"] = c_idx
            elif "увз" in group and "проголосовал" in sub:
                indices["voted_uvz"] = c_idx
            elif "19" in group and "планирует" in sub:
                indices["plan_u19"] = c_idx
            elif "19" in group and "открепился" in sub:
                indices["mark_u19"] = c_idx
            elif "19" in group and "проголосовал" in sub:
                indices["voted_u19"] = c_idx
            elif "уважительная" in combined or "причина" in combined:
                indices["absence"] = c_idx
        return indices

    def _has_mark(row: list, idx: int) -> bool:
        if idx is None or idx >= len(row):
            return False
        val = str(row[idx] or "").strip().lower()
        return val in ("+", "1", "да", "true", "v")

    def process_sheet(all_rows: list):
        nonlocal total_changed, total_rows, total_errors
        if len(all_rows) < 3:
            return

        header_idx = 0
        for idx, row in enumerate(all_rows):
            if any(
                "таб.№" in str(c).lower() or "табельный" in str(c).lower()
                for c in row
                if c
            ):
                header_idx = idx
                break

        header_rows = [all_rows[header_idx], all_rows[header_idx + 1]]
        indices = _get_col_indices(header_rows)

        if "tab" not in indices:
            total_errors += 1
            return

        updates = {}

        for row in all_rows[header_idx + 2 :]:
            if not any(row):
                continue
            total_rows += 1

            try:
                tab = _text(row[indices["tab"]])
                if not tab or not str(tab).isdigit():
                    total_errors += 1
                    continue
                tab = str(tab).zfill(7)

                if tab not in updates:
                    updates[tab] = {}

                u = updates[tab]

                if _has_mark(row, indices.get("plan_deg")):
                    u["method"] = DEG
                elif _has_mark(row, indices.get("plan_uik")):
                    u["method"] = UIK
                elif _has_mark(row, indices.get("plan_uvz")):
                    u["method"] = UVZ
                elif _has_mark(row, indices.get("plan_u19")):
                    u["method"] = UIK19

                if _has_mark(row, indices.get("mark_deg")):
                    u["mark_deg"] = True
                if _has_mark(row, indices.get("mark_uvz")):
                    u["mark_uvz"] = True
                if _has_mark(row, indices.get("detached")):
                    u["detached"] = True
                if _has_mark(row, indices.get("absence")):
                    u["absence"] = True

                if _has_mark(row, indices.get("voted_deg")):
                    u["voted"] = True
                    u["voted_method"] = DEG
                elif _has_mark(row, indices.get("voted_uik")):
                    u["voted"] = True
                    u["voted_method"] = UIK
                elif _has_mark(row, indices.get("voted_uvz")):
                    u["voted"] = True
                    u["voted_method"] = UVZ
                elif _has_mark(row, indices.get("voted_u19")):
                    u["voted"] = True
                    u["voted_method"] = UIK19

            except Exception:
                total_errors += 1
                continue

        if updates:
            tabs_list = list(updates.keys())
            existing = {
                emp.tab_number: emp
                for emp in Employee.objects.filter(tab_number__in=tabs_list)
            }

            to_update = []
            for tab, fields in updates.items():
                emp = existing.get(tab)
                if not emp:
                    total_errors += 1
                    continue

                changed = False
                for field, value in fields.items():
                    if getattr(emp, field) != value:
                        setattr(emp, field, value)
                        changed = True

                if changed:
                    to_update.append(emp)

            if to_update:
                Employee.objects.bulk_update(
                    to_update,
                    fields=list(set(f for u in updates.values() for f in u.keys())),
                    batch_size=BATCH,
                )
                total_changed += len(to_update)

    try:
        if upload.name.lower().endswith(".zip"):
            with zipfile.ZipFile(upload, "r") as zip_ref:
                for file_info in zip_ref.infolist():
                    if file_info.filename.endswith(
                        ".xlsx"
                    ) and not file_info.filename.startswith("__MACOSX"):
                        with zip_ref.open(file_info) as excel_file:
                            file_content = io.BytesIO(zip_ref.read(file_info.filename))
                            with _sheet(file_content) as rows:
                                process_sheet(list(rows))
        else:
            with _sheet(upload) as rows:
                process_sheet(list(rows))

    except zipfile.BadZipFile:
        raise ValueError("Ошибка: файл не является корректным zip-архивом")

    return total_changed, total_rows, total_errors
