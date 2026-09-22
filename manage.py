"""
Точка входа для управленческих команд Django.

Через этот файл идут все команды manage.py: migrate, runserver, test
и собственные команды (setup_groups, import_workshops).
"""

import os
import sys


def main():
    """
    Запускает управленческую команду из командной строки.

    Указывает Django, где лежит модуль настроек, и передаёт управление
    обработчику команд. Если Django не импортируется (не установлен или
    не активирован venv) — вместо голого ImportError выводит подсказку.
    """
    # setdefault, а не "=": значение из внешнего окружения имеет приоритет.
    # Это позволяет запускать проект с альтернативными настройками,
    # выставив DJANGO_SETTINGS_MODULE вручную.
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "elections.settings")
    try:
        from django.core.management import execute_from_command_line
    except ImportError as exc:
        # Импорт внутри try/except нужен только ради этой подсказки:
        # самая частая причина ошибки — забыл активировать .venv
        raise ImportError(
            "Couldn't import Django. Are you sure it's installed and "
            "available on your PYTHONPATH environment variable? Did you "
            "forget to activate a virtual environment?"
        ) from exc
    execute_from_command_line(sys.argv)


if __name__ == "__main__":
    main()
