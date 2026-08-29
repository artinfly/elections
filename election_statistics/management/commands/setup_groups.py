from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "Создаёт группу operator — ей разрешена загрузка файлов"

    def handle(self, *args, **kwargs):
        """
        Создаёт группу operator, если её ещё нет (get_or_create — команда идемпотентна,
        повторный запуск ничего не ломает). Пользователей в группу добавляют вручную
        через админку; сама группа даёт право заходить на страницу загрузки файлов.
        """
        Group.objects.get_or_create(name="operator")
        self.stdout.write("Группа operator создана")
