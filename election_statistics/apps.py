"""
Конфигурация приложения.

Описание:
    Содержит базовые настройки приложения, такие как его системное имя
    и человекочитаемое название для административной панели.
"""

from django.apps import AppConfig


class ElectionStatisticsConfig(AppConfig):
    """
    Конфигурация приложения election_statistics.
    Функциональных правок здесь не требуется: новые отчёты и поля модели
    не затрагивают настройки приложения.
    """
    default_auto_field = "django.db.models.BigAutoField"
    name = "election_statistics"

    def ready(self):
        import election_statistics.signals
