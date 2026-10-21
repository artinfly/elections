"""
Файл настроек проекта.

Описание:
    Содержит конфигурацию Django-приложения: подключение приложений,
    настройки базы данных, статики, безопасности и локализации.
    Пароли, ключ и адреса серверов в код не пишутся: они берутся из файла .env
    в корне проекта (в git не попадает, образец — .env.example).
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# ==============================================================================
# Базовые пути и безопасность
# ==============================================================================

# Корень проекта — папка, в которой лежит файл manage.py.
BASE_DIR = Path(__file__).resolve().parent.parent

# Переменные окружения из .env (уже заданные в системе имеют приоритет).
load_dotenv(BASE_DIR / ".env")

# Режим отладки. На сервере — False: подробные ошибки не показываются,
# статика раздаётся собранной (collectstatic).
DEBUG = os.getenv("DEBUG", "False").lower() == "true"

# Секретный ключ для подписи сессий и CSRF-токенов. На сервере обязателен
# свой ключ в .env; без него запуск с DEBUG=False остановится с ошибкой.
SECRET_KEY = (
    os.getenv("SECRET_KEY", "django-insecure-local-dev-key")
    if DEBUG
    else os.environ["SECRET_KEY"]
)

# Адреса, по которым открывают сайт (через запятую в .env).
ALLOWED_HOSTS = [
    host.strip() for host in os.getenv("ALLOWED_HOSTS", "").split(",") if host.strip()
]

# ==============================================================================
# Приложения и middleware
# ==============================================================================

# Список установленных приложений.
# Порядок важен: приложения, зависящие от других, должны быть ниже.
INSTALLED_APPS = [
    "core",
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
]

# Список middleware (прослоек), обрабатывающих запросы и ответы.
# Порядок обработки запроса: сверху вниз.
# Порядок обработки ответа: снизу вверх.
MIDDLEWARE = [
    # GZipMiddleware стоит первым в списке, чтобы при обработке ответа (снизу вверх)
    # он выполнялся последним и сжимал уже готовый контент.
    "django.middleware.gzip.GZipMiddleware",
    # SecurityMiddleware должен быть выше всех, кроме GZip, чтобы добавлять
    # заголовки безопасности до обработки другими модулями.
    "django.middleware.security.SecurityMiddleware",
    # WhiteNoiseMiddleware отвечает за раздачу статических файлов.
    # В режиме DEBUG берёт файлы напрямую из исходников.
    "whitenoise.middleware.WhiteNoiseMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    # Защита от межсайтовой подделки запросов (CSRF).
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    # Защита от кликджекинга (запрет встраивания в iframe).
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

# Путь к корневому файлу маршрутов.
ROOT_URLCONF = "config.urls"

# Настройки шаблонизатора.
TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        # Папки для поиска шаблонов (пусто, так как используются папки приложений).
        "DIRS": [],
        # Искать шаблоны внутри папок приложений (папка templates).
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

# WSGI-приложение для запуска на сервере.
WSGI_APPLICATION = "config.wsgi.application"

# ==============================================================================
# База данных
# ==============================================================================

# Подключение к PostgreSQL — значения из .env.
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.postgresql",
        "NAME": os.getenv("DB_NAME", "elections"),
        "USER": os.getenv("DB_USER", ""),
        "PASSWORD": os.getenv("DB_PASSWORD", ""),
        "HOST": os.getenv("DB_HOST", "localhost"),
        "PORT": os.getenv("DB_PORT", "5432"),
        "OPTIONS": {"client_encoding": "UTF8"},
    }
}

# ==============================================================================
# Аутентификация
# ==============================================================================

# Валидаторы паролей для обеспечения безопасности.
AUTH_PASSWORD_VALIDATORS = [
    {
        "NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"
    },
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]

# URL страницы входа (используется декоратором @login_required).
LOGIN_URL = "/login/"
# Куда перенаправлять пользователя после успешного входа.
LOGIN_REDIRECT_URL = "/"
# Куда перенаправлять пользователя после выхода.
LOGOUT_REDIRECT_URL = "/login/"

# ==============================================================================
# Локализация и время
# ==============================================================================

# Язык интерфейса.
LANGUAGE_CODE = "ru-ru"
# Часовой пояс проекта.
TIME_ZONE = "Asia/Yekaterinburg"
# Включить поддержку интернационализации.
USE_I18N = True
# Включить поддержку часовых поясов.
USE_TZ = True

# ==============================================================================
# Статика и загрузка файлов
# ==============================================================================

# Префикс для URL статических файлов.
STATIC_URL = "/static/"
# Папка, куда собираются статические файлы командой collectstatic.
STATIC_ROOT = BASE_DIR / "staticfiles"

# Настройки хранилищ файлов (новый синтаксис Django 4.2+).
STORAGES = {
    # Хранилище по умолчанию для пользовательских загрузок.
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    # Хранилище для статических файлов.
    "staticfiles": {
        "BACKEND": (
            # В DEBUG режиме используется простое хранилище без кэша.
            "django.contrib.staticfiles.storage.StaticFilesStorage"
            if DEBUG
            # В продакшене используется WhiteNoise со сжатием и хешами в именах.
            else "whitenoise.storage.CompressedManifestStaticFilesStorage"
        )
    },
}

# Флаги WhiteNoise для режима разработки.
# Позволяют раздавать статику из исходников и подхватывать изменения без перезапуска.
WHITENOISE_USE_FINDERS = DEBUG
WHITENOISE_AUTOREFRESH = DEBUG

# Максимальный размер файла (в байтах), который хранится в памяти при загрузке.
# Файлы крупнее 20 МБ пишутся во временный файл на диске.
FILE_UPLOAD_MAX_MEMORY_SIZE = 20 * 1024 * 1024

# ==============================================================================
# Прочее
# ==============================================================================

# Тип автоинкрементного первичного ключа по умолчанию для всех моделей.
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# Адрес кадрового API (админка: ФИО по табельному номеру, проверка увольнений).
HR_SERVICE_API_URL = os.getenv("HR_SERVICE_API_URL", "")
