"""
Маршруты приложения election_statistics.

Подключаются из корня проекта (elections/urls.py) через include.
Все маршруты имеют name — используется для redirect() в views и шаблонах.
"""

from django.urls import path

from . import views

urlpatterns = [
    # ==============================================================================
    # Аутентификация
    # ==============================================================================
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    # ==============================================================================
    # Основные страницы
    # ==============================================================================
    path("", views.method_page, name="method"),
    path("elections/", views.elections_page, name="elections"),
    path("upload/", views.upload_page, name="upload"),
    path("export/", views.export_page, name="export"),
    # ==============================================================================
    # Загрузка данных
    # ==============================================================================
    path("upload/base/", views.upload_base, name="upload_base"),
    path(
        "upload/voting-choices/",
        views.upload_voting_choices,
        name="upload_voting_choices",
    ),
    # ==============================================================================
    # Экспорт отчетов
    # ==============================================================================
    path("export/employees/", views.export_employees, name="export_employees"),
    path("export/summary/", views.export_summary, name="export_summary"),
    path(
        "export/summary-no-u19/",
        views.export_summary_no_19,
        name="export_summary_no_19",
    ),
    path("export/productions/", views.export_productions, name="export_productions"),
    path(
        "export/productions-methods/",
        views.export_production_methods,
        name="export_production_methods",
    ),
    path(
        "export/productions-methods-no-u19/",
        views.export_production_methods_no_19,
        name="export_production_methods_no_19",
    ),
    path("export/archive/", views.export_archive, name="export_archive"),
    path(
        "export/archive-methods/",
        views.export_method_archive,
        name="export_method_archive",
    ),
    path("export/custom/", views.export_custom_report, name="export_custom"),
    path(
        "export/custom-archive/",
        views.export_custom_archive,
        name="export_custom_archive",
    ),
    # ==============================================================================
    # API
    # ==============================================================================
    path("api/method/", views.api_method, name="api_method"),
    path("api/mark/", views.api_mark, name="api_mark"),
    path("api/voted/", views.api_voted, name="api_voted"),
    path("api/bulk-voted/", views.api_bulk_voted, name="api_bulk_voted"),
    path("api/uik-stats/", views.api_uik_stats, name="api_uik_stats"),
    path(
        "toggle-absence/<int:employee_id>/", views.toggle_absence, name="toggle_absence"
    ),
]
