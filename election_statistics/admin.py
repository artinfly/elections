"""
Модуль настройки административной панели Django.

Описание:
    Регистрирует модели Employee, EmployeeArchive и User в админке.
    Настраивает кастомные формы для управления пользователями с дополнительными
    полями из модели Profile (отчество, API-ключ, статус увольнения).
    Реализует интеграцию с внешним кадровым API для синхронизации данных.
"""

import requests
from django import forms
from django.conf import settings
from django.contrib import admin, messages
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.forms import UserChangeForm, UserCreationForm
from django.contrib.auth.models import User
from django.db import transaction
from django.http import JsonResponse
from django.urls import path
from django.utils import timezone

from .models import Employee, EmployeeArchive, Profile

# Сначала снимаем стандартную регистрацию модели User,
# чтобы переопределить её нашим кастомным классом.
admin.site.unregister(User)

API_PATH = settings.HR_SERVICE_API_URL
BULK_SYNC_LIMIT = 100


# ==============================================================================
# Формы для управления пользователями
# ==============================================================================


class CustomUserCreationForm(UserCreationForm):
    """
    Форма создания нового пользователя с дополнительными полями из Profile.
    """

    patronymic = forms.CharField(label="Отчество", max_length=255, required=False)
    api_key = forms.CharField(label="API-ключ", max_length=64, required=False)
    is_fired = forms.BooleanField(
        label="Уволен?", required=False, widget=forms.CheckboxInput
    )

    class Meta:
        model = User
        fields = ("username", "first_name", "last_name")


class CustomUserChangeForm(UserChangeForm):
    """
    Форма редактирования пользователя с дополнительными полями из Profile.
    """

    patronymic = forms.CharField(label="Отчество", max_length=255, required=False)
    api_key = forms.CharField(label="API-ключ", max_length=64, required=False)
    is_fired = forms.BooleanField(
        label="Уволен?", required=False, widget=forms.CheckboxInput
    )

    class Meta:
        model = User
        fields = "__all__"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance and self.instance.pk:
            profile, _ = Profile.objects.get_or_create(user=self.instance)
            self.fields["patronymic"].initial = profile.patronymic
            self.fields["api_key"].initial = profile.api_key
            self.fields["is_fired"].initial = profile.is_fired


# ==============================================================================
# Админка пользователей
# ==============================================================================


@admin.register(User)
class CustomUserAdmin(UserAdmin):
    """
    Кастомное отображение списка пользователей в админке.

    Описание:
        Оставляет только необходимые колонки для управления доступом.
        Сортирует пользователей по времени последнего входа для удобства
        мониторинга активности. Интегрируется с внешним кадровым API.
    """

    add_form = CustomUserCreationForm
    form = CustomUserChangeForm
    add_form_template = "admin/auth/user/add_form.html"
    change_form_template = "admin/auth/user/change_form.html"

    list_display = (
        "username",
        "get_full_name",
        "is_staff",
        "is_active",
        "last_login",
        "get_is_fired",
        "get_last_synced_at",
    )
    list_filter = (
        "is_staff",
        "is_superuser",
        "is_active",
        # Внимание: фильтр по дате последнего входа может быть ресурсоёмким
        # на очень больших базах пользователей.
        "last_login",
        "profile__is_fired",
    )

    ordering = ("-last_login", "username")
    search_fields = ("^username",)  # ^ означает поиск "начинается с", что быстрее

    actions = ["sync_with_external_api"]

    fieldsets = (
        (None, {"fields": ("username", "password")}),
        (
            "Персональная информация",
            {
                "fields": (
                    "last_name",
                    "first_name",
                    "patronymic",
                    "api_key",
                    "is_fired",
                )
            },
        ),
        (
            "Права доступа",
            {
                "fields": (
                    "is_active",
                    "is_staff",
                    "is_superuser",
                    "groups",
                    "user_permissions",
                )
            },
        ),
        ("Важные даты", {"fields": ("last_login", "date_joined")}),
    )

    add_fieldsets = (
        (
            None,
            {
                "classes": ("wide",),
                "fields": (
                    "username",
                    "last_name",
                    "first_name",
                    "patronymic",
                    "is_fired",
                    "api_key",
                    "password1",
                    "password2",
                ),
            },
        ),
    )

    def get_queryset(self, request):
        """Оптимизация: подгружаем Profile одним запросом."""
        return super().get_queryset(request).select_related("profile")

    @admin.display(description="ФИО", ordering="last_name")
    def get_full_name(self, obj):
        """Возвращает полное ФИО пользователя с учётом отчества из Profile."""
        patronymic = (
            getattr(obj.profile, "patronymic", "") if hasattr(obj, "profile") else ""
        )
        parts = [obj.last_name, obj.first_name, patronymic]
        return " ".join(p for p in parts if p)

    @admin.display(description="API ключ", ordering="profile__api_key")
    def get_api_key(self, obj):
        """Возвращает API-ключ из связанного Profile."""
        return getattr(obj.profile, "api_key", "")

    @admin.display(description="Уволен?", boolean=True, ordering="profile__is_fired")
    def get_is_fired(self, obj):
        """Возвращает статус увольнения из связанного Profile."""
        return bool(getattr(obj.profile, "is_fired", False))

    @admin.display(description="Дата синхронизации", ordering="profile__last_synced_at")
    def get_last_synced_at(self, obj):
        """Возвращает дату последней синхронизации из связанного Profile."""
        return getattr(obj.profile, "last_synced_at", "")

    def save_model(self, request, obj, form, change):
        """Сохраняет пользователя и синхронизирует дополнительные поля в Profile."""
        super().save_model(request, obj, form, change)
        Profile.objects.update_or_create(
            user=obj,
            defaults={
                "patronymic": form.cleaned_data.get("patronymic", ""),
                "api_key": form.cleaned_data.get("api_key", ""),
                "is_fired": form.cleaned_data.get("is_fired", ""),
            },
        )

    def get_urls(self):
        """Добавляет кастомный URL для получения данных из внешнего API."""
        custom_urls = [
            path(
                "fetch-external-data/<str:tab_number>/",
                self.admin_site.admin_view(self.fetch_external_data),
                name="auth_user_fetch_external_data",
            ),
        ]
        return custom_urls + super().get_urls()

    def fetch_external_data(self, request, tab_number):
        """
        Получает данные сотрудника из внешнего кадрового API по табельному номеру.

        Возвращает:
            JsonResponse с данными сотрудника или ошибкой.
        """
        api_key = getattr(getattr(request.user, "profile", None), "api_key", None)
        if not api_key:
            return JsonResponse(
                {"error": "У текущего пользователя не задан API ключ"}, status=400
            )

        url = f"{API_PATH}{tab_number}/"

        try:
            response = requests.get(url, headers={"X-API-Key": api_key}, timeout=5)
            response.raise_for_status()
        except requests.RequestException as e:
            return JsonResponse({"error": f"Ошибка обращения к API: {e}"}, status=502)

        try:
            data = response.json()
        except ValueError:
            return JsonResponse({"error": "Некорректный ответ от API"}, status=502)

        result = {
            "surname": data.get("surname", ""),
            "name": data.get("name", ""),
            "patronymic": data.get("patronymic", ""),
            "birth_date": data.get("birth_date", ""),
            "hire_date": data.get("hire_date", ""),
            "dismissal_date": data.get("dismissal_date", ""),
            "production": data.get("production", ""),
            "department": data.get("department", ""),
            "position": data.get("position", ""),
            "is_fired": data.get("is_fired", ""),
            "api_key": data.get("api_key", ""),
        }

        return JsonResponse(result)

    @admin.action(
        description=f"Синхронизация с сервисом персонала (До {BULK_SYNC_LIMIT} за раз)"
    )
    def sync_with_external_api(self, request, queryset):
        """
        Массовая синхронизация пользователей с внешним кадровым API.

        Описание:
            Обрабатывает до BULK_SYNC_LIMIT пользователей за раз, сортируя их
            по дате последней синхронизации (сначала давно не обновлявшиеся).
        """
        queryset = queryset.select_related("profile").order_by(
            "profile__last_synced_at"
        )
        total_selected = queryset.count()
        to_process = list(queryset[:BULK_SYNC_LIMIT])

        update_count = 0
        error_count = 0

        api_key = getattr(getattr(request.user, "profile", None), "api_key", None)
        if not api_key:
            self.message_user(request, "У вас не задан API-ключ!", level=messages.ERROR)
            return

        for user in to_process:
            profile, _ = Profile.objects.get_or_create(user=user)
            tab_number = user.username

            if not tab_number:
                profile.sync_error = "Не указано имя пользователя"
                profile.save(update_fields=["sync_error"])
                error_count += 1
                continue

            url = f"{API_PATH}{tab_number}/"
            try:
                response = requests.get(url, headers={"X-API-Key": api_key}, timeout=5)
                response.raise_for_status()
                data = response.json()
            except (requests.RequestException, ValueError) as e:
                profile.sync_error = str(e)
                profile.save(update_fields=["sync_error"])
                error_count += 1
                continue

            profile.user.first_name = data.get("name", profile.user.first_name)
            profile.user.last_name = data.get("surname", profile.user.last_name)
            profile.user.save(update_fields=["first_name", "last_name"])

            profile.patronymic = data.get("patronymic", profile.patronymic)
            profile.is_fired = data.get("is_fired", profile.is_fired)
            profile.api_key = data.get("api_key", profile.api_key)
            profile.last_synced_at = timezone.now()
            profile.sync_error = ""
            profile.save()

            update_count += 1

        skipped = total_selected - len(to_process)
        msg = f"Обновлено: {update_count}. Ошибок: {error_count}."
        if skipped > 0:
            msg += (
                f" Не обработано (Превышен лимит {BULK_SYNC_LIMIT} за раз): {skipped}"
            )
        self.message_user(request, msg)


# ==============================================================================
# Админка сотрудников
# ==============================================================================


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
        "absence",
        "mark_uvz",
        "mark_deg",
        "proxy_vote",
    )

    # Вычисляемые поля (ФИО, читаемые способы) доступны только для просмотра
    # в детальной карточке сотрудника.
    readonly_fields = ("fio", "method_label", "voted_method_label")

    # Поля, редактируемые прямо в списке (обязательно должны быть в list_display).
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
        "proxy_vote",
    )

    actions = ["check_and_archive_fired", "unmark_wrong_absence"]

    @admin.action(
        description="Архивация уволенных (Может занять продолжительное время...)"
    )
    def check_and_archive_fired(self, request, queryset):
        """
        Проверяет статус увольнения через внешний API и архивирует уволенных сотрудников.

        Описание:
            Для каждого выбранного сотрудника делает запрос к кадровому API.
            Если сотрудник уволен, создаёт запись в EmployeeArchive и удаляет из Employee.
        """
        api_key = getattr(getattr(request.user, "profile", None), "api_key", None)
        if not api_key:
            self.message_user(request, "У вас не задан API-ключ", level=messages.ERROR)
            return

        total_selected = queryset.count()
        to_process = list(queryset)

        archived_count = 0
        error_count = 0
        checked_count = 0

        for employee in to_process:
            url = f"{API_PATH}{employee.tab_number}/"
            try:
                response = requests.get(url, headers={"X-API-Key": api_key}, timeout=5)
                response.raise_for_status()
                data = response.json()
            except (requests.RequestException, ValueError) as e:
                self.message_user(
                    request,
                    f"Ошибка API для {employee.tab_number}: {e}",
                    level=messages.WARNING,
                )
                error_count += 1
                continue

            checked_count += 1
            is_fired = bool(data.get("is_fired", False))

            if not is_fired:
                continue

            try:
                with transaction.atomic():
                    EmployeeArchive.objects.create(
                        tab_number=employee.tab_number,
                        department=employee.department,
                        production=employee.production,
                        service=employee.service,
                        surname=employee.surname,
                        name=employee.name,
                        patronymic=employee.patronymic,
                        position=employee.position,
                        category=employee.category,
                        birth_date=employee.birth_date,
                        region=employee.region,
                        city=employee.city,
                        street=employee.street,
                        house=employee.house,
                        uik=employee.uik,
                        uik_address=employee.uik_address,
                        district=employee.district,
                        okrug=employee.okrug,
                        method=employee.method,
                        voted=employee.voted,
                        voted_method=employee.voted_method,
                        voted_at=employee.voted_at,
                        detached=employee.detached,
                        not_going=employee.not_going,
                        mark_uvz=employee.mark_uvz,
                        mark_deg=employee.mark_deg,
                        absence=employee.absence,
                        proxy_vote=employee.proxy_vote,
                    )
                    employee.delete()
                archived_count += 1
            except Exception as e:
                self.message_user(
                    request,
                    f"Ошибка при архивации {employee.tab_number}: {e}",
                    level=messages.ERROR,
                )
                error_count += 1

        skipped = total_selected - len(to_process)
        # ИСПРАВЛЕНО: Было msg = {...} (множество), теперь f-строка
        msg = (
            f"Проверено {checked_count}. "
            f"Архивировано (уволены): {archived_count}. "
            f"Ошибок: {error_count}."
        )
        if skipped > 0:
            msg += f" Не обработано (лимит {BULK_SYNC_LIMIT} за раз): {skipped}."
        self.message_user(request, msg)

    @admin.action(description="Снять отметку об УП")
    def unmark_wrong_absence(self, request, queryset):
        """
        Снимает отметку 'Отсутствие по УП' у сотрудников, у которых выбран способ голосования.

        Описание:
            Логика: если сотрудник выбрал способ голосования, он не может отсутствовать по УП.
        """
        emp_list = [
            employee.tab_number
            for employee in queryset
            if employee.method.strip() and employee.absence
        ]

        if emp_list:
            Employee.objects.filter(tab_number__in=emp_list).update(absence=False)

        self.message_user(request, f"Снято отметок: {len(emp_list)}")


# ==============================================================================
# Админка архива сотрудников
# ==============================================================================


@admin.register(EmployeeArchive)
class EmployeeArchiveAdmin(admin.ModelAdmin):
    """
    Админка архива уволенных сотрудников.

    Описание:
        Предоставляет доступ только для просмотра. Создание и редактирование
        записей запрещено (только через action 'check_and_archive_fired').
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
        "absence",
        "mark_uvz",
        "mark_deg",
        "proxy_vote",
    )

    readonly_fields = ("fio", "method_label", "voted_method_label")

    list_editable = ("detached", "not_going", "absence")

    list_filter = (
        "method",
        "voted",
        "voted_method",
        "department",
        "detached",
        "not_going",
        "absence",
        "mark_uvz",
        "mark_deg",
        "proxy_vote",
    )

    def has_add_permission(self, request):
        """Запрещает создание записей через админку."""
        return False

    def has_change_permission(self, request, obj=None):
        """Запрещает редактирование записей через админку."""
        return False

    # Поиск "начинается с" (^) для оптимизации запросов к БД.
    # Позволяет быстро находить сотрудника по началу табельного номера или фамилии.
    search_fields = ("^tab_number", "^surname", "^name", "^patronymic")
    ordering = ("surname", "name", "patronymic")
