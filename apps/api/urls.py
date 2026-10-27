"""Адреса JSON API (подключены в config/urls.py под /api/; описание — docs/api.md).

Действия на страницах сайта (/api/method/, /api/voted/ и др.) — в core/urls.py.
"""

from django.urls import path

from . import views

app_name = "api"

urlpatterns = [
    # Данные страниц
    path("", views.method_page, name="method_page"),
    path("elections/", views.elections_page, name="elections_page"),
    path("upload/", views.upload_page, name="upload_page"),
    path("export/", views.export_page, name="export_page"),
    # Вход и выход
    path("csrf/", views.csrf, name="csrf"),
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    # Загрузка файлов
    path("upload/base/", views.upload_base, name="upload_base"),
    path("upload/voting-choices/", views.upload_voting_choices,
         name="upload_voting_choices"),
    path("upload/turnout-hq/", views.upload_turnout_hq, name="upload_turnout_hq"),
    path("upload/custom-report/", views.upload_custom_report,
         name="upload_custom_report"),
    path("upload/responsible/", views.upload_responsible_marks,
         name="upload_responsible_marks"),
    # Выгрузки
    path("export/summary/", views.export_summary, name="export_summary"),
    path("export/summary-no-u19/", views.export_summary_no_19,
         name="export_summary_no_19"),
    path("export/productions/", views.export_productions, name="export_productions"),
    path("export/productions-methods/", views.export_production_methods,
         name="export_production_methods"),
    path("export/productions-methods-no-u19/", views.export_production_methods_no_19,
         name="export_production_methods_no_19"),
    path("export/employees/", views.export_employees, name="export_employees"),
    path("export/uvz-statement-not-voted/", views.export_uvz_statement_not_voted,
         name="export_uvz_statement_not_voted"),
    path("export/responsible-template/", views.export_responsible_template,
         name="export_responsible_template"),
    path("export/archive/", views.export_archive, name="export_archive"),
    path("export/custom-archive/", views.export_custom_archive,
         name="export_custom_archive"),
    path("export/archive-methods/", views.export_method_archive,
         name="export_method_archive"),
    path("export/custom/", views.export_custom, name="export_custom"),
]  # fmt: skip
