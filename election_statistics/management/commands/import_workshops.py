import getpass

import psycopg2
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from election_statistics.models import Employee
from election_statistics.services import padded_number

# Запрос к внешней базе: пары "номер цеха -> название производства".
# Читается только справочник, данные работников внешней базы не трогаем.
QUERY = """
    select cw."number", cp."name"
    from core_workshop cw
    join core_production cp on cp.id = cw.production_id
    where cp."name" is not null
"""


class Command(BaseCommand):
    help = "Переносит справочник цех -> производство из базы"

    def add_arguments(self, parser):
        # Параметры подключения к внешней базе; пароль не передаётся аргументом,
        # а запрашивается интерактивно, чтобы не светиться в истории команд
        parser.add_argument("--host", required=True)
        parser.add_argument("--dbname", required=True)
        parser.add_argument("--user", required=True)
        parser.add_argument("--port", default="5432")

    def handle(self, *args, **options):
        # --- Шаг 1. Читаем справочник из внешней базы ---
        password = getpass.getpass("Пароль: ")
        source = None
        try:
            source = psycopg2.connect(
                host=options["host"],
                port=options["port"],
                dbname=options["dbname"],
                user=options["user"],
                password=password,
            )
            with source.cursor() as cursor:
                cursor.execute(QUERY)
                fetched = cursor.fetchall()
        except psycopg2.Error as error:
            raise CommandError(f"Не удалось прочитать базу: {error}")
        finally:
            # Соединение закрываем в любом случае, даже при ошибке
            if source is not None:
                source.close()

        # --- Шаг 2. Собираем словарь {цех: производство} с нормализацией номеров ---
        # Номера цехов приводятся к единому виду (padded_number: "7" -> "007"),
        # пустые значения отбрасываются
        pairs = {}
        for number, production in fetched:
            number = str(number or "").strip()
            production = str(production or "").strip()
            if number and production:
                pairs[padded_number(number)] = production

        # Пустой справочник — ошибка, чтобы молча не стереть разметку
        if not pairs:
            raise CommandError("Запрос не вернул ни одной пары цех - производство")

        # --- Шаг 3. Группируем цеха нашей базы по производствам ---
        departments = list(
            Employee.objects.exclude(department="")
            .values_list("department", flat=True)
            .distinct()
        )
        by_production = {}
        for department in departments:
            production = pairs.get(padded_number(department))
            if production:
                by_production.setdefault(production, []).append(department)

        # --- Шаг 4. Атомарно переразмечаем поле production у работников ---
        # Сначала сбрасываем старую разметку, затем ставим новую по справочнику.
        # Всё в одной транзакции: при сбое база не останется полустёртой
        with transaction.atomic():
            Employee.objects.exclude(production="").update(production="")
            marked = 0
            for production, group in by_production.items():
                marked += Employee.objects.filter(department__in=group).update(
                    production=production
                )

        # --- Шаг 5. Итоговая статистика для оператора ---
        self.stdout.write(f"Строк получено: {len(fetched)}")
        self.stdout.write(f"Цехов в справочнике: {len(pairs)}")
        self.stdout.write(f"Производств: {len(by_production)}")
        self.stdout.write(f"Работников размечено: {marked}")

        # Цеха, которые есть у работников, но отсутствуют в справочнике,
        # выводим красным — их нужно добавить во внешнюю базу или поправить запрос
        used = {padded_number(value) for value in departments}
        missing = sorted(used - set(pairs))
        if missing:
            self.stdout.write(
                self.style.ERROR(
                    f"Нет в справочнике, но есть у работников ({len(missing)}): "
                    + ", ".join(missing)
                )
            )
        else:
            self.stdout.write(self.style.SUCCESS("Все цеха базы нашлись в справочнике"))

        # Цеха из справочника, по которым у нас нет людей, — просто информация
        unused = sorted(set(pairs) - used)
        if unused:
            self.stdout.write(
                f"В справочнике есть, но людей нет ({len(unused)}): "
                + ", ".join(unused)
            )
