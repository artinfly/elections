from django.apps import AppConfig


class ApiConfig(AppConfig):
    """JSON API для отдельного фронтенда: данные страниц, вход, загрузки и выгрузки."""

    name = "apps.api"
    label = "api"
    verbose_name = "API для фронтенда"
