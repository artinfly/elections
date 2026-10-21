"""Конфигурация приложения core — учёт голосования работников."""

from django.apps import AppConfig


class CoreConfig(AppConfig):
    """
    Настройки приложения.

    Папка называется core, но внутреннее имя (label) оставлено прежним —
    election_statistics. От него зависят имена таблиц в базе
    (election_statistics_employee и т.д.), записи о миграциях и права
    пользователей, поэтому переименование папки базу не затрагивает.
    """

    name = "core"
    label = "election_statistics"
    verbose_name = "Учёт голосования"
    default_auto_field = "django.db.models.BigAutoField"

    def ready(self):
        """Подключает обработчики сигналов (создание профиля пользователя)."""
        # Импорт ради побочного эффекта: при импорте регистрируются @receiver
        from . import signals  # noqa: F401
