from django.contrib import admin

from .models import Employee


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    list_display = ("tab_number", "fio", "department", "uik", "method", "voted")
    list_filter = ("method", "voted", "department")
    search_fields = ("tab_number", "surname", "name", "patronymic")
    ordering = ("surname", "name")
