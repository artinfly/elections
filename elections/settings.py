"""
Настройки проекта.

"""

from pathlib import Path

# ==============================================================================
# База
# ==============================================================================

# Корень проекта — папка, в которой лежит manage.py
BASE_DIR = Path(__file__).resolve().parent.parent

# Ключ для dev; на сервере должен быть собственный
SECRET_KEY = "django-insecure-%r9ot_pgbi2=6$=i&yt0m@l7i4t4*1=vso8%a7678!lsry4k2k"
DEBUG = True
ALLOWED_HOSTS = ["localhost", "127.0.0.1", "testserver"]

# ==============================================================================
# Приложения и middleware
# ==============================================================================

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "election_statistics",
]

MIDDLEWARE = [
    # Стоит первым: ответ обрабатывается снизу вверх,
    # поэтому сжимается уже готовая страница
    "django.middleware.gzip.GZipMiddleware",
    "django.middleware.security.SecurityMiddleware",
    # Отдаёт статику; в DEBUG берёт файлы из исходников через finders
    "whitenoise.middleware.WhiteNoiseMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "elections.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
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

WSGI_APPLICATION = "elections.wsgi.application"

# ==============================================================================
# База данных
# ==============================================================================

# Dev-база; на сервере меняются имя, пользователь, пароль и хост
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.postgresql",
        "NAME": "elections",
        "USER": "root",
        "PASSWORD": "root",
        "HOST": "localhost",
        "PORT": "5432",
        # Кириллица в базе без сюрпризов с кодировкой
        "OPTIONS": {"client_encoding": "UTF8"},
    }
}

# ==============================================================================
# Аутентификация
# ==============================================================================

AUTH_PASSWORD_VALIDATORS = [
    {
        "NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"
    },
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]

# Куда попадёт пользователь после входа/выхода,
# если в ссылке не указан next
LOGIN_URL = "/login/"
LOGIN_REDIRECT_URL = "/"
LOGOUT_REDIRECT_URL = "/login/"

# ==============================================================================
# Локализация и время
# ==============================================================================

LANGUAGE_CODE = "ru-ru"
TIME_ZONE = "Asia/Yekaterinburg"
USE_I18N = True
USE_TZ = True

# ==============================================================================
# Статика и загрузка файлов
# ==============================================================================

STATIC_URL = "/static/"
STATIC_ROOT = BASE_DIR / "staticfiles"

# В DEBUG — обычное хранилище, собирать ничего не нужно.
# В проде — whitenoise со сжатием и хешами в именах,
# поэтому предварительно запускается collectstatic
STORAGES = {
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    "staticfiles": {
        "BACKEND": (
            "django.contrib.staticfiles.storage.StaticFilesStorage"
            if DEBUG
            else "whitenoise.storage.CompressedManifestStaticFilesStorage"
        )
    },
}

# Dev-флаги whitenoise: статика из исходников и подхват изменений без collectstatic
WHITENOISE_USE_FINDERS = DEBUG
WHITENOISE_AUTOREFRESH = DEBUG

# Файл до этого размера держится в памяти, крупнее — пишется во временный
FILE_UPLOAD_MAX_MEMORY_SIZE = 20 * 1024 * 1024

# ==============================================================================
# Прочее
# ==============================================================================

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"
