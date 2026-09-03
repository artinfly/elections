"""
Файл настроек проекта.

Описание:
    Содержит конфигурацию Django-приложения: подключение приложений,
    настройки базы данных, статики, безопасности и локализации.
    Внимание: значения, помеченные как "только для разработки",
    должны быть заменены при развертывании на боевом сервере.
"""

from pathlib import Path

# ==============================================================================
# Базовые пути и безопасность
# ==============================================================================

# Корень проекта — папка, в которой лежит файл manage.py.
BASE_DIR = Path(__file__).resolve().parent.parent

# Секретный ключ для криптографической подписи (сессии, CSRF-токены).
# ВНИМАНИЕ: В продакшене ключ должен быть уникальным, сложным и храниться
# в переменных окружения, а не в коде. Текущее значение — только для разработки.
SECRET_KEY = "django-insecure-%r9ot_pgbi2=6$=i&yt0m@l7i4t4*1=vso8%a7678!lsry4k2k"

# Режим отладки. ВНИМАНИЕ: В продакшене должен быть False.
# При True показываются подробные ошибки и раздаются статические файлы без кэша.
DEBUG = True

# Список хостов, с которых разрешено принимать запросы.
# В продакшене нужно указать доменное имя сервера.
ALLOWED_HOSTS = ["localhost", "127.0.0.1", "testserver"]

# ==============================================================================
# Приложения и middleware
# ==============================================================================

# Список установленных приложений.
# Порядок важен: приложения, зависящие от других, должны быть ниже.
INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "election_statistics",
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
ROOT_URLCONF = "elections.urls"

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
WSGI_APPLICATION = "elections.wsgi.application"

# ==============================================================================
# База данных
# ==============================================================================

# Настройки подключения к PostgreSQL.
# ВНИМАНИЕ: В продакшене учетные данные (USER, PASSWORD) и хост
# должны быть изменены на безопасные.
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.postgresql",
        "NAME": "elections",
        "USER": "root",
        "PASSWORD": "root",
        "HOST": "localhost",
        "PORT": "5432",
        # Явное указание кодировки для корректной работы с кириллицей.
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
