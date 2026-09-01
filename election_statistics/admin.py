from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User

from .models import Employee

# Сначала снимаем стандартную регистрацию модели User,
# чтобы переопределить её нашим кастомным классом
admin.site.unregister(User)


@admin.register(User)
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
    search_fields = ("^username",)  # ^ означает поиск "начинается с", что быстрее


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    """
    Админка сотрудников: список с основными колонками, фильтры по способам,
    явке и подразделениям, поиск по табельному номеру и ФИО.
    Отметки «Открепился» и «Не пойдет» проставляются галочками
    прямо в списке (list_editable) для быстрого заполнения сводного отчёта.
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

    # Делаем вычисляемые поля доступными для просмотра в детальной карточке
    readonly_fields = ("fio", "method_label", "voted_method_label")

    # Поля, редактируемые прямо в списке (обязательно должны быть в list_display)
    list_editable = ("detached", "not_going")

    list_filter = (
        "method",
        "voted",
        "voted_method",
        "department",
        "detached",
        "not_going",
    )

    # ^ означает поиск "начинается с", что значительно ускоряет запросы к БД
    search_fields = ("^tab_number", "^surname", "^name", "^patronymic")
    ordering = ("surname", "name", "patronymic")
