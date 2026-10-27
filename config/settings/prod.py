"""Настройки для рабочего сервера: отладка выключена, адреса сайта обязательны.

Перед запуском: manage.py collectstatic (статику раздаёт whitenoise из staticfiles/).
"""

from .base import *  # noqa: F403
from .base import env_bool, env_list

DEBUG = False

# Адреса, по которым открывается сайт (например, имя сервера), — только из .env
ALLOWED_HOSTS = env_list("ALLOWED_HOSTS")

# Сжатая статика с версиями в именах файлов (браузер не держит старый CSS)
STORAGES = {
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    "staticfiles": {
        "BACKEND": "whitenoise.storage.CompressedManifestStaticFilesStorage"
    },
}

# Сайт за nginx: адрес и протокол берутся из заголовков прокси
USE_X_FORWARDED_HOST = env_bool("USE_X_FORWARDED_HOST")
if env_bool("SECURE_PROXY_SSL_HEADER"):
    SECURE_PROXY_SSL_HEADER = ("HTTP_X_FORWARDED_PROTO", "https")

# Cookie входа и защиты форм — только по защищённому соединению, если сайт на HTTPS
SESSION_COOKIE_SECURE = env_bool("HTTPS")
CSRF_COOKIE_SECURE = SESSION_COOKIE_SECURE

SECURE_CONTENT_TYPE_NOSNIFF = True
X_FRAME_OPTIONS = "DENY"
# CSRF_COOKIE_HTTPONLY не включаем: фронтенд API читает токен из cookie csrftoken
