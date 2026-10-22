"""Вход, выход и права ролей."""

from django.test import TestCase
from django.urls import reverse

from core.models import Employee

from .helpers import make_employee, make_user, xlsx


class AuthTests(TestCase):
    """Страницы закрыты без входа; выход — только кнопкой (POST)."""

    def test_pages_require_login(self):
        for name in ("method", "elections", "upload", "export"):
            response = self.client.get(reverse(name))
            self.assertRedirects(response, f"/login/?next={reverse(name)}")

    def test_login_with_password(self):
        make_user("user")
        response = self.client.post(
            reverse("login"), {"username": "user", "password": "pass"}
        )
        self.assertRedirects(response, reverse("method"))

    def test_wrong_password_shows_error(self):
        make_user("user")
        response = self.client.post(
            reverse("login"), {"username": "user", "password": "x"}
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.context["error"])

    def test_logout_only_by_post(self):
        self.client.force_login(make_user("user"))
        self.client.get(reverse("logout"))
        self.assertEqual(self.client.get(reverse("method")).status_code, 200)
        self.client.post(reverse("logout"))
        self.assertEqual(self.client.get(reverse("method")).status_code, 302)


class RoleTests(TestCase):
    """Загрузка — только оператору; наблюдатели не меняют данные."""

    def setUp(self):
        self.employee = make_employee()

    def test_uploads_need_operator(self):
        self.client.force_login(make_user("user"))
        uploads = (
            "upload_base",
            "upload_turnout_hq",
            "upload_custom_report",
            "upload_responsible_marks",
        )
        for name in uploads:
            response = self.client.post(
                reverse(name), {"file": xlsx([["Таб№"], ["1"]])}
            )
            # Не открываем страницу загрузки: она показывает сообщение и убирает его
            self.assertRedirects(
                response, reverse("upload"), fetch_redirect_response=False
            )
            self.assertIn("только оператор", self.client.session["msg"])
        self.assertEqual(Employee.objects.count(), 1)

    def test_operator_can_upload(self):
        self.client.force_login(make_user("op", "operator"))
        upload = xlsx(
            [
                ["Таб№", "Подразделение", "Фамилия", "Имя"],
                ["2000001", "5", "Новиков", "Олег"],
            ]
        )
        response = self.client.post(reverse("upload_base"), {"file": upload})
        self.assertRedirects(response, reverse("upload"))
        self.assertTrue(Employee.objects.filter(tab_number="2000001").exists())

    def test_viewer_cannot_change_data(self):
        self.client.force_login(make_user("viewer", "viewer"))
        response = self.client.post(
            reverse("api_method"),
            {"id": self.employee.id, "method": "deg"},
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 403)

    def test_uik_uvz_viewer_sees_only_turnout_page(self):
        self.client.force_login(make_user("uvz", "uik_uvz_viewer"))
        self.assertRedirects(self.client.get(reverse("method")), reverse("elections"))
        self.assertRedirects(self.client.get(reverse("upload")), reverse("elections"))
        self.assertEqual(self.client.get(reverse("elections")).status_code, 200)
