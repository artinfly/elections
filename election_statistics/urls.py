from django.urls import path

from . import views

urlpatterns = [
    path("", views.method_page, name="method"),
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("elections/", views.elections_page, name="elections"),
    path("upload/", views.upload_page, name="upload"),
    path("upload/base/", views.upload_base, name="upload_base"),
    path("export/", views.export_page, name="export"),
    path("export/employees/", views.export_employees, name="export_employees"),
    path("export/archive/", views.export_archive, name="export_archive"),
    path("api/method/", views.api_method, name="api_method"),
    path("api/voted/", views.api_voted, name="api_voted"),
    path("api/bulk-voted/", views.api_bulk_voted, name="api_bulk_voted"),
    path("api/uik-stats/", views.api_uik_stats, name="api_uik_stats"),
]
