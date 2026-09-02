"""
Корневая карта маршрутов проекта.

Все рабочие маршруты описаны внутри приложения election_statistics;
здесь только подключается админка и корень делегируется приложению.
"""

from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path("admin/", admin.site.urls),
    # Все маршруты приложения — в election_statistics/urls.py
    path("", include("election_statistics.urls")),
]
