from django.urls import path

from . import views

urlpatterns = [
    path("", views.method_page, name="method"),
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("elections/", views.elections_page, name="elections"),
    path("upload/", views.upload_page, name="upload"),
    path("upload/base/", views.upload_base, name="upload_base"),
    path("upload/methods/", views.upload_methods, name="upload_methods"),
    path("upload/voted/", views.upload_voted, name="upload_voted"),
    path("export/", views.export_page, name="export"),
    path("api/method/", views.api_method, name="api_method"),
    path("api/voted/", views.api_voted, name="api_voted"),
    path("api/bulk-voted/", views.api_bulk_voted, name="api_bulk_voted"),
]
