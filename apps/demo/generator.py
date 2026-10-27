"""Вымышленные сотрудники и голосование для разработки.

Основной список и выбор способов голосования загружаются через тот же импорт
Excel, что на странице «Загрузка» (import_base, import_voting_choices), — так
демо-данные заодно проверяют настоящий импорт. Явка отмечается функцией сайта
set_turnout. Все люди, адреса и участки вымышленные.
"""

import io
import random
from datetime import date, timedelta

from django.db import transaction
from openpyxl import Workbook

from core.models import DEG, UIK, UIK19, UVZ, Employee
from core.services.excel import COLUMNS
from core.services.importing import (
    CHOICE_COLUMNS,
    import_base,
    import_voting_choices,
    set_turnout,
)

SURNAMES = ["Иванов", "Смирнов", "Кузнецов", "Попов", "Васильев", "Петров",
            "Соколов", "Михайлов", "Новиков", "Фёдоров", "Морозов", "Волков",
            "Алексеев", "Лебедев", "Семёнов", "Егоров", "Павлов", "Козлов"]  # fmt: skip
MALE_NAMES = ["Александр", "Сергей", "Дмитрий", "Андрей", "Алексей", "Иван",
              "Михаил", "Николай", "Евгений", "Владимир", "Павел", "Олег"]  # fmt: skip
FEMALE_NAMES = ["Елена", "Ольга", "Наталья", "Татьяна", "Ирина", "Светлана",
                "Анна", "Мария", "Юлия", "Екатерина"]  # fmt: skip
POSITIONS = ["Слесарь-ремонтник", "Токарь", "Инженер-технолог", "Мастер участка",
             "Сварщик", "Кладовщик", "Экономист", "Контролёр ОТК", "Электромонтёр"]  # fmt: skip
STREETS = ["ул. Заводская", "ул. Садовая", "ул. Мира", "ул. Лесная", "пр. Строителей"]

# Цех -> (производство, служба); номера цехов трёхзначные, как в основной базе
WORKSHOPS = {
    f"{number:03d}": (f"Производство №{1 + number % 4}", f"Служба №{1 + number % 3}")
    for number in range(1, 25)
}
OKRUGS = ["19", "20", "21", ""]

# Выбор способа: обычный участок чаще всего, часть — без выбора
METHOD_WEIGHTS = {UIK: 50, UVZ: 20, DEG: 15, UIK19: 5, "": 10}
TITLE_BY_METHOD = {method: title for title, method in CHOICE_COLUMNS.items()}


def _xlsx(rows):
    """Файл Excel в памяти (как загруженный через форму)."""
    book = Workbook()
    for row in rows:
        book.active.append(row)
    content = io.BytesIO()
    book.save(content)
    content.seek(0)
    content.name = "demo.xlsx"
    return content


class DemoData:
    """Вымышленные строки файлов; seed — одинаковый результат при каждом запуске."""

    def __init__(self, people=600, seed=1):
        self.random = random.Random(seed)
        self.people = [self._person(number) for number in range(1, people + 1)]

    def _person(self, number):
        rnd = self.random
        male = rnd.random() < 0.6
        surname = rnd.choice(SURNAMES)
        father = rnd.choice(MALE_NAMES)
        okrug = rnd.choice(OKRUGS)
        uik = f"{1000 + rnd.randint(1, 40)}" if okrug else ""
        return {
            "department": rnd.choice(list(WORKSHOPS)),
            "tab_number": f"{100000 + number:07d}",
            "surname": surname if male else surname + "а",
            "name": rnd.choice(MALE_NAMES if male else FEMALE_NAMES),
            "patronymic": father + ("ович" if male else "овна"),
            "position": rnd.choice(POSITIONS),
            "category": rnd.choice(["Рабочий", "Специалист", "Руководитель"]),
            "birth_date": date(1960, 1, 1) + timedelta(days=rnd.randint(0, 15_000)),
            "region": "Демо-область",
            "city": "г. Энск",
            "street": rnd.choice(STREETS),
            "house": str(rnd.randint(1, 120)),
            "uik": uik,
            "uik_address": f"г. Энск, школа №{uik[-2:]}" if uik else "",
            "district": rnd.choice(["Северный", "Южный", "Центральный"]),
            "okrug": okrug,
        }

    def base_file(self):
        fields = list(COLUMNS.values())
        rows = [list(COLUMNS)] + [[p[f] for f in fields] for p in self.people]
        return _xlsx(rows)

    def choices_file(self):
        """Отчёт штаба: «1» в колонке выбранного способа."""
        methods = list(METHOD_WEIGHTS)
        titles = list(CHOICE_COLUMNS)
        rows = [["Отчёт штаба (демо-данные)"], ["Таб.№", "ФИО", *titles]]
        for p in self.people:
            method = self.random.choices(methods, weights=METHOD_WEIGHTS.values())[0]
            marks = [
                1 if title == TITLE_BY_METHOD.get(method) else None for title in titles
            ]
            rows.append([p["tab_number"], f"{p['surname']} {p['name']}", *marks])
        return _xlsx(rows)


@transaction.atomic
def seed(people=600, seed_value=1):
    """Заменяет всех сотрудников вымышленными. Возвращает итоги по шагам."""
    data = DemoData(people, seed_value)
    rnd = data.random
    Employee.objects.all().delete()

    created, _updated, _total = import_base(data.base_file())
    for workshop, (production, service) in WORKSHOPS.items():
        Employee.objects.filter(department=workshop).update(
            production=production, service=service
        )
    chosen, _rows, _errors = import_voting_choices(data.choices_file())

    # Регистрация на ДЭГ и УВЗ у части выбравших, отсутствие по УП у немногих
    Employee.objects.filter(method=DEG, pk__in=_sample(rnd, DEG, 0.7)).update(
        mark_deg=True
    )
    Employee.objects.filter(method=UVZ, pk__in=_sample(rnd, UVZ, 0.8)).update(
        mark_uvz=True
    )
    Employee.objects.filter(method="", pk__in=_sample(rnd, "", 0.3)).update(
        absence=True
    )

    # Явка: проголосовала примерно половина выбравших способ
    voters = [pk for pk in Employee.objects.exclude(method="").values_list("pk", flat=True)
              if rnd.random() < 0.55]  # fmt: skip
    voted = set_turnout(Employee.objects.filter(pk__in=voters), True)
    return {"сотрудников": created, "выбрали способ": chosen, "проголосовали": voted}


def _sample(rnd, method, share):
    pks = Employee.objects.filter(method=method).values_list("pk", flat=True)
    return [pk for pk in pks if rnd.random() < share]
