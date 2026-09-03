"""
ASGI-точка входа проекта elections.

Сейчас не используется: проект полностью синхронный и отдаётся по WSGI
(wsgi.py). Файл понадобится только при переходе на ASGI-сервер
(uvicorn, daphne) или появлении асинхронных фич — например, WebSocket.
"""

import os

from django.core.asgi import get_asgi_application

# setdefault, как в manage.py: значение из внешнего окружения имеет приоритет
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "elections.settings")

application = get_asgi_application()
