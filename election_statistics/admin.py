"""
Модуль административной панели.

Описание:
    Настраивает отображение моделей в админке. Переопределяет стандартный
    интерфейс пользователей и добавляет кастомные настройки для сотрудников,
    включая редактируемые поля в списке и фильтры.
"""

from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User

from .models import Employee

# Отменяем стандартную регистрацию модели User, чтобы переопределить её
# нашим кастомным классом ниже.
admin.site.unregister(User)


@admin.register(User)
class CustomUserAdmin(UserAdmin):
    """
    Кастомное отображение списка пользователей в админке.

    Описание:
        Оставляет только необходимые колонки для управления доступом.
        Сортирует пользователей по времени последнего входа для удобства
        мониторинга активности.
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
        # Внимание: фильтр по дате последнего входа может быть ресурсоемким
        # на очень больших базах пользователей.
        "last_login",
    )
    ordering = ("-last_login", "username")
    # Поиск "начинается с" (^) работает быстрее, чем поиск по вхождению.
    search_fields = ("^username",)


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    """
    Админка сотрудников.

    Описание:
        Выводит список с основными колонками и фильтрами по способам,
        явке и подразделениям. Позволяет быстро редактировать служебные
        отметки (открепился, не пойдет, отсутствие по УП) прямо в списке
        без перехода в карточку сотрудника.
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
        # Добавлены поля для контроля статуса УП и регистраций.
        "absence",
        "mark_uvz",
        "mark_deg",
    )

    # Вычисляемые поля (ФИО, читаемые способы) доступны только для просмотра
    # в детальной карточке сотрудника.
    readonly_fields = ("fio", "method_label", "voted_method_label")

    # Поля, редактируемые прямо в списке (обязательно должны быть в list_display).
    # Добавлено absence для управления отметкой УП из админки.
    list_editable = ("detached", "not_going", "absence")

    list_filter = (
        "method",
        "voted",
        "voted_method",
        # Внимание: при большом количестве уникальных цехов этот фильтр
        # может замедлять загрузку страницы админки.
        "department",
        "detached",
        "not_going",
        "absence",
        "mark_uvz",
        "mark_deg",
    )

    # Поиск "начинается с" (^) для оптимизации запросов к БД.
    # Позволяет быстро находить сотрудника по началу табельного номера или фамилии.
    search_fields = ("^tab_number", "^surname", "^name", "^patronymic")
    ordering = ("surname", "name", "patronymic")
