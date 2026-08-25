from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User

from .models import Employee

class CustomUserAdmin(UserAdmin):
    list_display = (
        'username',
        'is_staff',
        'is_active',
        'last_login',
    )

    list_filter = (
        'is_staff',
        'is_superuser',
        'is_active',
        'last_login',
    )

    ordering = ('-last_login', 'username')
    search_fields = ('username',)

admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)

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
