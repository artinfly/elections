"""Настройки для разработки: режим отладки, статика прямо из папок проекта."""

from .base import *  # noqa: F403
from .base import INSTALLED_APPS, env_list

DEBUG = True

ALLOWED_HOSTS = env_list("ALLOWED_HOSTS", "localhost,127.0.0.1")

# Генератор вымышленных данных (manage.py seed_demo) — только для разработки
INSTALLED_APPS = [*INSTALLED_APPS, "apps.demo"]

# Статика отдаётся без collectstatic и подхватывает правки без перезапуска
STORAGES = {
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    "staticfiles": {"BACKEND": "django.contrib.staticfiles.storage.StaticFilesStorage"},
}
WHITENOISE_USE_FINDERS = True
WHITENOISE_AUTOREFRESH = True
