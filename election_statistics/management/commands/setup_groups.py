"""
Создаёт группу operator (идемпотентно, повторный запуск безопасен).
Пользователей в группу добавляют вручную через админку.
"""

from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "Создаёт группу operator — ей разрешена загрузка файлов."

    def handle(self, *args, **kwargs):
        _, created = Group.objects.get_or_create(name="operator")
        self.stdout.write(
            "Группа operator создана" if created else "Группа operator уже существует"
        )
