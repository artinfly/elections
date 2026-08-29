from django.apps import AppConfig


class ElectionStatisticsConfig(AppConfig):
    """
    Конфигурация приложения election_statistics.
    Функциональных правок здесь не требуется: новые отчёты и поля модели
    не затрагивают настройки приложения.
    """

    name = "election_statistics"
