"""
Модуль импорта и обновления данных сотрудников из Excel-файлов.

Описание:
    Содержит логику чтения выгрузок из Excel и записи данных в базу.
    Поддерживает три сценария:
    1. Импорт основной базы сотрудников (ФИО, адреса, участки).
    2. Импорт отчета штаба с выбранными способами голосования.
    3. Массовая простановка явки (из списка табельных номеров).
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
    Разбирает строки файла в словарь по табельному номеру.

    Описание:
        Проходит по строкам листа и собирает данные в словарь, где ключом
        является табельный номер. Если в файле есть дубликаты строк,
        последняя запись перезаписывает предыдущие.

    Аргументы:
        rows: итератор строк листа.
        positions: словарь {имя_колонки: индекс} из функции _header.

    Возвращает:
        dict: {табельный_номер: {поле: значение, ...}}.
    """
    parsed = {}
    for row in rows:
        if not any(row):
            continue
        values = {}
        for name, index in positions.items():
            field = COLUMNS[name]
            cell = row[index] if index < len(row) else None
            # Даты парсим отдельно, остальные поля приводим к строке.
            values[field] = _date(cell) if field == "birth_date" else _text(cell)

        tab = values.pop("tab_number")
        # Приводим числовые табельные номера к единому формату (дополняем нулями).
        if tab.isdigit():
            tab = tab.zfill(7)
        if tab:
            parsed[tab] = values
    return parsed


def _known_rows(tabs: list, fields: list) -> dict:
    """
    Подтягивает существующих сотрудников из БД чанками.

    Описание:
        Чтобы не загружать всю таблицу сотрудников в оперативную память,
        запрос к базе разбивается на пакеты по 2000 табельных номеров.

    Аргументы:
        tabs: список табельных номеров из загружаемого файла.
        fields: список полей, которые нужно достать для сравнения.

    Возвращает:
        dict: {табельный_номер: {"pk": ID, поля...}}.
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
    Импорт основной базы сотрудников из кадровой выгрузки.

    Описание:
        Создает новые записи и обновляет изменившиеся данные у существующих.
        Важная особенность: поля, связанные с голосованием (способ, явка, производство),
        НЕ стираются при повторной загрузке, так как их нет в исходном файле.

    Аргументы:
        upload: файл или поток с xlsx.

    Возвращает:
        tuple: (создано, обновлено, всего строк).

    Исключения:
        ValueError: битый файл или не найдена обязательная колонка "Таб№".
    """
    with _sheet(upload) as rows:
        positions = _header(rows, COLUMNS)
        if "Таб№" not in positions:
            raise ValueError("Ошибка: не найдена колонка 'Таб№'")
        parsed = _rows_by_tab(rows, positions)

    if not parsed:
        return 0, 0, 0

    # Список полей для проверки изменений (все, кроме табельного номера).
    fields = [COLUMNS[name] for name in positions if COLUMNS[name] != "tab_number"]
    known = _known_rows(parsed, fields)

    fresh, stale = [], []
    for tab, values in parsed.items():
        current = known.get(tab)
        if current is None:
            # Сотрудника нет в базе — готовим к созданию.
            fresh.append(Employee(tab_number=tab, **values))
        elif any(current[field] != values[field] for field in fields):
            # Данные изменились — готовим к обновлению (используем существующий PK).
            stale.append(Employee(pk=current["pk"], tab_number=tab, **values))

    # Пакетные операции для минимизации нагрузки на БД.
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

    Описание:
        При отметке явки место голосования (voted_method) автоматически копируется
        из плана (method). При снятии отметки поле voted_method очищается.
        Сам план (method) не изменяется.

    Аргументы:
        queryset: QuerySet сотрудников для отметки.
        voted: True — отметить явку, False — снять.

    Возвращает:
        int: число изменённых строк.
    """
    return queryset.update(
        voted=voted,
        voted_at=timezone.now() if voted else None,
        voted_method=F("method") if voted else "",
    )


def mark_voted(tabs: list, voted: bool = True) -> tuple[int, int]:
    """
    Отмечает явку по списку табельных номеров (используется API).

    Аргументы:
        tabs: список табельных номеров.
        voted: True — отметить, False — снять.

    Возвращает:
        tuple: (изменено, не найдено в базе).
    """
    tabs = {t for t in tabs if t}
    if not tabs:
        return 0, 0
    found = Employee.objects.filter(tab_number__in=tabs)
    missing = len(tabs) - found.count()
    return set_turnout(found, voted), missing


# ==============================================================================
# Импорт отчётов штаба (Способы голосования и Явка)
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

    Описание:
        Читает файл, где способы голосования отмечены единицами ("1") в соответствующих
        колонках. Обновляет поле method у существующих сотрудников.
        Сотрудники, которых нет в базе, НЕ создаются: без department, okrug и uik
        они сломали бы отчёты по цехам и округам — такая строка считается ошибкой.

    Оптимизация:
        Для избежания N+1 запросов все табельные номера сначала парсятся из файла,
        затем одним запросом извлекаются из БД, и обновление происходит пакетно (bulk_update).

    Аргументы:
        upload: файл или поток с xlsx.

    Возвращает:
        tuple: (обновлено, всего строк, ошибок/пропусков).

    Исключения:
        ValueError: битый или пустой файл, не найдены обязательные колонки.
    """
    with _sheet(upload) as rows:
        all_rows = list(rows)

    if not all_rows:
        raise ValueError("Ошибка: файл пустой")

    # Поиск строки заголовка (регистр и точное положение не важны).
    header_row_idx = 0
    for idx, row in enumerate(all_rows):
        cells = [str(cell or "").strip() for cell in row if cell is not None]
        if "Таб.№" in cells or "ФИО" in cells:
            header_row_idx = idx
            break

    if header_row_idx >= len(all_rows) - 1:
        raise ValueError("Ошибка: не найдена строка с данными")

    # Сопоставление колонок.
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

    # --- Этап 1: Парсинг файла в память ---
    file_data = {}  # {табельный_номер: выбранный_способ}
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

            # ИСПРАВЛЕНО: добавлены скобки .isdigit(), иначе zfill ломал буквенные табельные.
            if tab_number.isdigit():
                tab_number = tab_number.zfill(7)

            # Способ определяется отметкой "1". Две единицы в строке — ошибка формата.
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
                # Строки без способа учитываются в общем количестве, но не в ошибках.
                continue

            file_data[tab_number] = selected_method

        except Exception:
            parse_errors += 1
            continue

    if not file_data:
        return 0, total, parse_errors

    # --- Этап 2: Запрос к БД и подготовка к обновлению ---
    # Достаем всех существующих сотрудников одним запросом.
    existing_employees = {
        emp.tab_number: emp
        for emp in Employee.objects.filter(tab_number__in=file_data.keys())
    }

    employees_to_update = []
    not_found_errors = 0

    for tab_number, selected_method in file_data.items():
        employee = existing_employees.get(tab_number)
        if employee is None:
            # Не создаём "обрывков" без цеха и округа.
            not_found_errors += 1
            continue

        # Обновляем только если способ действительно изменился.
        if employee.method != selected_method:
            employee.method = selected_method
            employees_to_update.append(employee)

    # --- Этап 3: Пакетное сохранение ---
    if employees_to_update:
        Employee.objects.bulk_update(
            employees_to_update, fields=["method"], batch_size=BATCH
        )

    updated = len(employees_to_update)
    errors = parse_errors + not_found_errors

    return updated, total, errors


def import_turnout(upload: Any) -> tuple[int, int, int]:
    """
    Импорт отметок явки из файла штаба.

    Описание:
        Ожидает файл, где в первой строке (ячейка A1) находится заголовок,
        содержащий слово "Табельный" или "Таб", а в последующих строках
        в первом столбце указаны табельные номера проголосовавших сотрудников.
        Проставляет явку (voted=True) и время отметки.
        Сотрудники без заранее выбранного способа голосования пропускаются,
        чтобы не нарушать бизнес-логику (нельзя проголосовать, не выбрав способ).

    Аргументы:
        upload: файл или поток с xlsx.

    Возвращает:
        tuple: (количество отмеченных, всего строк в файле, количество ошибок/пропусков).

    Исключения:
        ValueError: если файл пустой или отсутствует ожидаемый заголовок.
    """
    with _sheet(upload) as rows:
        all_rows = list(rows)

    if not all_rows:
        raise ValueError("Ошибка: файл пустой")

    # Проверяем заголовок в первой ячейке первой строки.
    header = str(all_rows[0][0] or "").strip().lower()
    if "табель" not in header and "таб" not in header:
        raise ValueError("Ошибка: в А1 ожидается заголовок 'Табельный'")

    tabs = []
    for row in all_rows[1:]:  # Пропускаем строку заголовка
        if not row:
            continue
        tab = _text(row[0])
        if tab and tab.isdigit():
            tabs.append(tab.zfill(7))

    if not tabs:
        return 0, 0, 0

    total_rows = len(tabs)

    # Ищем всех сотрудников из списка в базе.
    found = Employee.objects.filter(tab_number__in=tabs)
    missing_in_db = total_rows - found.count()

    # Фильтруем тех, у кого уже выбран способ голосования.
    # Бизнес-правило: нельзя отметить явку, если не выбран способ.
    valid_found = found.exclude(method="")
    skipped_no_method = found.filter(method="").count()

    # Массово проставляем явку одним SQL-запросом.
    changed = set_turnout(valid_found, voted=True)

    # Ошибками считаем тех, кого нет в базе, и тех, у кого не был выбран способ.
    errors = missing_in_db + skipped_no_method

    return changed, total_rows, errors
