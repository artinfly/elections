from django.contrib import admin

from .models import Employee


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    list_display = (
        "tab_number",
        "fio",
        "department",
        "uik",
        "method",
        "voted",
        "voted_method",
    )
    list_filter = ("method", "voted", "voted_method", "department")
    search_fields = ("tab_number", "surname", "name", "patronymic")
    ordering = ("surname", "name")
