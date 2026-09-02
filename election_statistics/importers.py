"""
Модуль импорта и обновления данных сотрудников из Excel-файлов.

Два формата входа: основная база сотрудников (import_base) и отчёт штаба
с выбранными способами голосования (import_voting_choices).
Плюс массовая простановка явки (set_turnout, mark_voted).
"""

from typing import Any, Iterator

from django.db.models import F
from django.utils import timezone

from .helpers import BATCH, COLUMNS, _date, _header, _sheet, _text
from .models import Employee

# ==============================================================================
# Импорт основной базы
# ==============================================================================


def _rows_by_tab(rows: Iterator, positions: dict) -> dict:
    """
    Разбирает строки файла в словарь {табельный номер: значения полей}.

    Аргументы:
        rows: итератор строк листа.
        positions: {имя_колонки: индекс} из _header.

    Возвращает:
        Словарь по табельному номеру; дубликаты строк в файле перезаписываются.
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
        if tab:
            parsed[tab] = values
    return parsed


def _known_rows(tabs: list, fields: list) -> dict:
    """
    Подтягивает существующих сотрудников из БД чанками по 2000,
    чтобы не держать в памяти всю таблицу сразу.

    Аргументы:
        tabs: список табельных номеров из файла.
        fields: какие поля достать для сравнения.

    Возвращает:
        Словарь {табельный номер: {"pk": ..., поля...}}.
    """
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
    """
    Импорт основной базы сотрудников.

    Повторная загрузка обновляет изменившиеся строки и не трогает
    способ голосования, явку и производство: этих колонок нет в файле.

    Аргументы:
        upload: файл или поток с xlsx.

    Возвращает:
        Кортеж (создано, обновлено, всего строк).

    Исключения:
        ValueError: битый файл или не найдена колонка "Таб№".
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
            # В UPDATE попадают только реально изменившиеся строки
            stale.append(Employee(pk=current["pk"], tab_number=tab, **values))

    if fresh:
        Employee.objects.bulk_create(fresh, batch_size=BATCH)
    if stale:
        Employee.objects.bulk_update(stale, fields, batch_size=BATCH)

    return len(fresh), len(stale), len(parsed)


# ==============================================================================
# Отметки явки
# ==============================================================================


def set_turnout(queryset: Any, voted: bool = True) -> int:
    """
    Массово проставляет явку одним UPDATE-запросом.

    При отметке явки место голосования копируется из плана (method),
    при снятии — очищается; сам план не меняется.

    Аргументы:
        queryset: QuerySet сотрудников для отметки.
        voted: True — отметить явку, False — снять.

    Возвращает:
        Число изменённых строк.
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
        tabs: список табельных номеров (пустые значения игнорируются).
        voted: True — отметить, False — снять.

    Возвращает:
        Кортеж (изменено, не найдено в базе).
    """
    tabs = {t for t in tabs if t}
    if not tabs:
        return 0, 0
    found = Employee.objects.filter(tab_number__in=tabs)
    missing = len(tabs) - found.count()
    return set_turnout(found, voted), missing


# ==============================================================================
# Импорт нового формата: "Сводная таблица для отчета штабу"
# ==============================================================================

# Маркеры колонок отчёта штаба: способ голосования определяется отметкой "1"
NEW_FORMAT_COLUMNS = {
    "Таб.№": "tab_number",
    "ФИО": "fio_raw",
    "Очно на своем избирательном участке": "uik",
    "При помощи дистанционного электронного голосования ДЭГ": "deg",
    "Очно на временном избирательном участке на территории предприятия": "uvz",
    "Очно на избирательном участке 19 округа": "u19",
}


def import_voting_choices(upload: Any) -> tuple[int, int, int]:
    """
    Импорт способов голосования из отчёта штаба.

    Обновляет поле method у существующих сотрудников. Сотрудники, которых
    нет в базе, НЕ создаются: без department, okrug и uik они сломали бы
    отчёты по цехам и округам — такая строка считается ошибкой.

    Аргументы:
        upload: файл или поток с xlsx.

    Возвращает:
        Кортеж (обновлено, всего строк, ошибок/пропусков).

    Исключения:
        ValueError: битый или пустой файл, не найдены обязательные колонки.
    """
    # _sheet гарантирует закрытие книги даже при ошибке
    with _sheet(upload) as rows:
        all_rows = list(rows)

    if not all_rows:
        raise ValueError("Ошибка: файл пустой")

    # Строка заголовка ищется по маркерам, её положение в файле не фиксировано
    header_row_idx = 0
    for idx, row in enumerate(all_rows):
        cells = [str(cell or "").strip() for cell in row if cell is not None]
        if "Таб.№" in cells or "ФИО" in cells:
            header_row_idx = idx
            break

    if header_row_idx >= len(all_rows) - 1:
        raise ValueError("Ошибка: не найдена строка с данными")

    # Заголовок может быть частью более длинного текста — ищем по вхождению
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

    updated = 0
    total = 0
    errors = 0

    for row in all_rows[header_row_idx + 1 :]:
        if not any(row):
            continue

        total += 1

        try:
            tab_col = col_indices["Таб.№"]
            tab_number = _text(row[tab_col] if tab_col < len(row) else None)
            if not tab_number:
                continue

            # Способ определяется отметкой "1"; две единицы в строке — ошибка
            selected_method = ""
            for method_code, col_idx in method_cols.items():
                if col_idx is not None and col_idx < len(row):
                    if _text(row[col_idx]) == "1":
                        if selected_method:
                            errors += 1
                            selected_method = ""
                            break
                        selected_method = method_code

            # Строки без способа учитываются в пропусках, но не как ошибки
            if not selected_method:
                continue

            try:
                employee = Employee.objects.get(tab_number=tab_number)
            except Employee.DoesNotExist:
                # Не создаём "обрывков" без цеха и округа — см. docstring
                errors += 1
                continue

            employee.method = selected_method
            employee.save(update_fields=["method"])
            updated += 1

        except Exception:
            # Одна битая строка не должна останавливать весь импорт
            errors += 1
            continue

    return updated, total, errors
