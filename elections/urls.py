"""Корневая карта маршрутов проекта."""

from django.contrib import admin
from django.urls import include, path

admin.site.site_header = "Учёт голосования — Администрирование"
admin.site.site_title = "Учёт голосования — Портал администратора"
admin.site.index_title = "Управление данными"

urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("election_statistics.urls")),
]
