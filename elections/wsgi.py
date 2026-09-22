"""
WSGI-точка входа проекта elections.

Модульная переменная ``application`` — то, что вызывает продакшен-сервер
(gunicorn, uWSGI) при обработке каждого запроса.
Разработочный сервер (runserver) этот файл не использует.
"""

import os

from django.core.wsgi import get_wsgi_application

# setdefault, как в manage.py: значение из внешнего окружения имеет приоритет
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "elections.settings")

application = get_wsgi_application()
