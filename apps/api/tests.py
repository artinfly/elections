import json

from django.test import TestCase

from core.models import DEG, UIK, UVZ, Employee
from core.tests.helpers import make_employee, make_user, xlsx


class ApiAuthTests(TestCase):
    def login(self, username, password="pass"):
        return self.client.post(
            "/api/login/",
            json.dumps({"username": username, "password": password}),
            content_type="application/json",
        )

    def test_csrf_token(self):
        response = self.client.get("/api/csrf/")
        self.assertTrue(response.json()["csrfToken"])
        self.assertIn("csrftoken", response.cookies)

    def test_login_roles_and_logout(self):
        make_user("op", "operator")
        response = self.login("op")
        self.assertEqual(response.json()["user"]["role"], "operator")
        self.assertEqual(self.login("op").json()["msg"], "Пользователь уже авторизован")
        self.assertEqual(self.client.post("/api/logout/").status_code, 200)
        self.assertEqual(self.client.post("/api/logout/").status_code, 401)

        make_user("watch", "viewer")
        self.assertEqual(self.login("watch").json()["user"]["role"], "viewer")

    def test_wrong_password_and_bad_json(self):
        make_user("op", "operator")
        response = self.login("op", "wrong")
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.json()["msg"], "Неверный логин или пароль")
        response = self.client.post("/api/login/", "{", content_type="application/json")
        self.assertEqual(response.status_code, 400)

    def test_anonymous_gets_401(self):
        for url in ("/api/", "/api/elections/", "/api/export/summary/"):
            self.assertEqual(self.client.get(url).status_code, 401, url)


class ApiPagesTests(TestCase):
    def setUp(self):
        make_employee(surname="Петров", name="Пётр", patronymic="", method=UVZ)
        make_employee(surname="Сидоров", name="Сидор", method=DEG)

    def test_method_page(self):
        self.client.force_login(make_user("op", "operator"))
        data = self.client.get("/api/", {"q": "Петров"}).json()
        self.assertEqual(data["pagination"]["total_count"], 1)
        self.assertEqual(data["results"][0]["fio"], "Петров Пётр")
        data = self.client.get("/api/", {"page": "999"}).json()
        self.assertEqual(data["pagination"]["current_page"], 1)

    def test_uvz_viewer(self):
        self.client.force_login(make_user("uvz", "uik_uvz_viewer"))
        self.assertEqual(self.client.get("/api/").json()["url"], "/elections")
        data = self.client.get("/api/elections/").json()
        self.assertEqual([r["fio"] for r in data["results"]], ["Петров Пётр"])
        self.assertNotIn("voted_method", data["results"][0])

    def test_export_page(self):
        self.client.force_login(make_user("op", "operator"))
        data = self.client.get("/api/export/").json()
        self.assertEqual(data["counts"]["method"]["total"], 2)
        self.assertTrue(data["is_operator"])


class ApiUploadExportTests(TestCase):
    def test_uploads_need_operator(self):
        self.client.force_login(make_user("user"))
        for url in ("/api/upload/base/", "/api/upload/turnout-hq/"):
            response = self.client.post(url, {"file": xlsx([["x"]])})
            self.assertEqual(response.status_code, 403, url)

    def test_wrong_extension(self):
        self.client.force_login(make_user("op", "operator"))
        response = self.client.post("/api/upload/turnout-hq/", {"file": xlsx([["x"]])})
        self.assertEqual(response.status_code, 400)
        self.assertIn(".zip", response.json()["msg"])

    def test_voting_choices(self):
        first = make_employee()
        second = make_employee()
        self.client.force_login(make_user("op", "operator"))
        header = [
            "Таб.№",
            "ФИО",
            "Очно на своем избирательном участке",
            "При помощи дистанционного электронного голосования ДЭГ",
        ]
        rows = [
            ["Отчёт штаба"],
            header,
            [first.tab_number, "А", 1, None],
            [second.tab_number, "Б", 1, 1],  # две отметки — ошибка
            ["9999999", "В", None, 1],  # нет в базе
        ]
        response = self.client.post("/api/upload/voting-choices/", {"file": xlsx(rows)})
        self.assertEqual(
            response.json()["details"], {"total": 3, "updated": 1, "errors_count": 2}
        )
        self.assertEqual(Employee.objects.get(pk=first.pk).method, UIK)
        self.assertEqual(Employee.objects.get(pk=second.pk).method, "")

    def test_exports(self):
        make_employee(method=UIK)
        self.client.force_login(make_user("op", "operator"))
        response = self.client.get("/api/export/summary/")
        self.assertEqual(response["Content-Type"].split(";")[0][-5:], "sheet")
        response = self.client.get("/api/export/archive/")
        self.assertEqual(response["Content-Type"], "application/zip")
        Employee.objects.all().delete()
        self.assertEqual(self.client.get("/api/export/archive/").status_code, 404)
