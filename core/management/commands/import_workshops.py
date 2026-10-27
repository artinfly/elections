"""
Команда import_workshops: проставляет сотрудникам производство по номеру цеха.

Справочник «цех — производство» читается из внешней базы PostgreSQL
(только чтение). Пароль к ней запрашивается при запуске.
"""

import getpass

import psycopg
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from core.models import Employee
from core.services.excel import padded_number

# Запрос к внешней базе: пары "номер цеха -> название производства".
# Читается только справочник, данные работников внешней базы не трогаем.
QUERY = """
    select cw."number", cp."name"
    from core_workshop cw
    join core_production cp on cp.id = cw.production_id
    where cp."name" is not null
"""


class Command(BaseCommand):
    """Перенос справочника «цех — производство» из внешней базы."""

    help = "Переносит справочник цех -> производство из базы"

    def add_arguments(self, parser):
        """Параметры подключения к внешней базе (без пароля)."""
        # Параметры подключения к внешней базе; пароль не передаётся аргументом,
        # а запрашивается интерактивно, чтобы не светиться в истории команд
        parser.add_argument("--host", required=True)
        parser.add_argument("--dbname", required=True)
        parser.add_argument("--user", required=True)
        parser.add_argument("--port", default="5432")

    def handle(self, *args, **options):
        """Читает справочник и атомарно переразмечает поле production у сотрудников."""
        # --- Шаг 1. Читаем справочник из внешней базы ---
        password = getpass.getpass("Пароль: ")
        # Соединение закрывается само при выходе из with, даже при ошибке
        try:
            with psycopg.connect(
                host=options["host"],
                port=options["port"],
                dbname=options["dbname"],
                user=options["user"],
                password=password,
            ) as source:
                fetched = source.execute(QUERY).fetchall()
        except psycopg.Error as error:
            raise CommandError(f"Не удалось прочитать базу: {error}") from error

        # --- Шаг 2. Собираем словарь {цех: производство} с нормализацией номеров ---
        # Номера цехов приводятся к единому виду (padded_number: "7" -> "007"),
        # пустые значения отбрасываются
        pairs = {}
        for raw_number, raw_production in fetched:
            number = str(raw_number or "").strip()
            production = str(raw_production or "").strip()
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
