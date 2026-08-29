from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User

from .models import Employee


class CustomUserAdmin(UserAdmin):
    """
    Кастомное отображение списка пользователей в админке:
    только нужные колонки и сортировка по времени последнего входа.
    """

    list_display = (
        "username",
        "is_staff",
        "is_active",
        "last_login",
    )

    list_filter = (
        "is_staff",
        "is_superuser",
        "is_active",
        "last_login",
    )

    ordering = ("-last_login", "username")
    search_fields = ("username",)


# Подменяем стандартную админку пользователей на кастомную
admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    """
    Админка сотрудников: список с основными колонками, фильтры по способам,
    явке и подразделениям, поиск по табельному номеру и ФИО.
    Новые отметки «Открепился» и «Не пойдет» проставляются галочками
    прямо в списке (list_editable) — так оператор быстро заполняет данные
    для сводного отчёта без изменения основных страниц.
    """

    list_display = (
        "tab_number",
        "fio",
        "department",
        "uik",
        "method",
        "voted",
        "voted_method",
        "detached",
        "not_going",
    )
    # Поля, редактируемые прямо в списке: появляются чекбоксы,
    # изменения сохраняются кнопкой "Сохранить" внизу страницы.
    # Поле из list_editable обязано присутствовать в list_display.
    list_editable = ("detached", "not_going")
    list_filter = (
        "method",
        "voted",
        "voted_method",
        "department",
        "detached",
        "not_going",
    )
    search_fields = ("tab_number", "surname", "name", "patronymic")
    ordering = ("surname", "name")
