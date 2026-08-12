from datetime import datetime

import openpyxl
from django.utils import timezone

from .models import DEG, METHOD_LABELS, UIK, UVZ, Employee

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


def _open(upload):
    try:
        book = openpyxl.load_workbook(upload, read_only=True, data_only=True)
    except Exception:
        raise ValueError(BAD_FORMAT)
    return book.active.iter_rows(values_only=True)


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


def import_base(upload):
    rows = _open(upload)
    positions = _header(rows, COLUMNS)
    if "Таб№" not in positions:
        raise ValueError(BAD_FORMAT)

    created = updated = 0
    for row in rows:
        if not any(row):
            continue
        values = {}
        for name, index in positions.items():
            field = COLUMNS[name]
            cell = row[index] if index < len(row) else None
            values[field] = _date(cell) if field == "birth_date" else _text(cell)
        tab = values.pop("tab_number")
        if not tab:
            continue
        _, is_new = Employee.objects.update_or_create(tab_number=tab, defaults=values)
        created += is_new
        updated += not is_new
    return created, updated


def norm_method(value):
    text = _text(value).upper().replace(" ", "").replace("-", "")
    if not text:
        return ""
    if "ДЭГ" in text or "ДЕГ" in text:
        return DEG
    if "УВЗ" in text:
        return UVZ
    if "УИК" in text:
        return UIK
    return ""


def import_methods(upload):
    rows = _open(upload)
    positions = _header(rows, ["Таб№", "Способ"])
    if len(positions) < 2:
        raise ValueError(BAD_FORMAT)

    changed = skipped = 0
    for row in rows:
        if not any(row):
            continue
        tab = _text(row[positions["Таб№"]])
        method = norm_method(row[positions["Способ"]])
        if not tab or not method:
            skipped += 1
            continue
        updated = Employee.objects.filter(tab_number=tab).update(method=method)
        changed += updated
        skipped += not updated
    return changed, skipped


def read_tab_column(upload):
    rows = _open(upload)
    index = _header(rows, ["Таб№"])["Таб№"]
    return [_text(row[index]) for row in rows if any(row) and _text(row[index])]


def parse_tabs(text):
    parts = text.replace(",", " ").replace(";", " ").split()
    return [p.strip() for p in parts if p.strip()]


def mark_voted(tabs, voted=True):
    tabs = {t for t in tabs if t}
    if not tabs:
        return 0, 0
    found = Employee.objects.filter(tab_number__in=tabs)
    missing = len(tabs) - found.count()
    changed = found.update(voted=voted, voted_at=timezone.now() if voted else None)
    return changed, missing


def export_xlsx():
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Сотрудники"
    sheet.append(list(COLUMNS) + ["Способ", "Проголосовал"])
    fields = [COLUMNS[name] for name in COLUMNS]
    for person in Employee.objects.all().iterator(chunk_size=2000):
        row = [getattr(person, f) for f in fields]
        row.append(METHOD_LABELS.get(person.method, ""))
        row.append("да" if person.voted else "нет")
        sheet.append(row)
    return book
