"""manage.py seed_demo — заменить сотрудников вымышленными (только для разработки)."""

from django.core.management.base import BaseCommand

from apps.demo.generator import seed


class Command(BaseCommand):
    help = "Вымышленные сотрудники, выбор способов и явка (заменяют всех сотрудников)"

    def add_arguments(self, parser):
        parser.add_argument("--people", type=int, default=600, help="число сотрудников")
        parser.add_argument("--seed", type=int, default=1, help="повторяемый результат")

    def handle(self, *args, **options):
        result = seed(options["people"], options["seed"])
        for name, value in result.items():
            self.stdout.write(f"{name}: {value}")
        self.stdout.write(self.style.SUCCESS("Демо-данные загружены"))
