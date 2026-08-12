from pathlib import Path

from django.contrib.auth.models import Group, User
from django.test import TestCase

from .models import DEG, Employee
from .services import import_base

TEMPLATE = Path(r"C:\Users\smiar\Desktop\excel\Книга1.xlsx")


class AccessTests(TestCase):
    def setUp(self):
        self.viewer = User.objects.create_user("viewer", password="x")
        self.operator = User.objects.create_user("operator", password="x")
        self.operator.groups.add(Group.objects.create(name="operator"))

    def test_anonymous_redirected_to_login(self):
        for url in ["/", "/elections/", "/upload/", "/export/"]:
            response = self.client.get(url)
            self.assertEqual(response.status_code, 302, url)
            self.assertIn("/login/", response["Location"], url)

    def test_logged_in_sees_pages(self):
        self.client.force_login(self.viewer)
        for url in ["/", "/elections/", "/export/"]:
            self.assertEqual(self.client.get(url).status_code, 200, url)

    def test_upload_hidden_from_plain_user(self):
        self.client.force_login(self.viewer)
        page = self.client.get("/upload/")
        self.assertContains(page, "Нет доступа")
        posted = self.client.post("/upload/base/", {})
        self.assertEqual(posted.status_code, 403)

    def test_operator_may_open_upload(self):
        self.client.force_login(self.operator)
        self.assertContains(self.client.get("/upload/"), "База сотрудников")

    def test_api_requires_login(self):
        self.assertEqual(
            self.client.post(
                "/api/method/", "{}", content_type="application/json"
            ).status_code,
            302,
        )


class ImportTests(TestCase):
    def test_template_columns_are_parsed(self):
        if not TEMPLATE.exists():
            self.skipTest("нет файла-образца")
        created, updated = import_base(TEMPLATE)
        self.assertEqual(updated, 0)
        self.assertTrue(created)

        person = Employee.objects.get(tab_number="0848103")
        self.assertEqual(person.department, "97")
        self.assertEqual(person.surname, "Амбарян")
        self.assertEqual(person.position, "Врач-хирург")
        self.assertEqual(str(person.birth_date), "1975-11-17")
        self.assertEqual(person.uik, "2632")
        self.assertEqual(person.okrug, "21")

    def test_second_import_keeps_method_and_voted(self):
        if not TEMPLATE.exists():
            self.skipTest("нет файла-образца")
        import_base(TEMPLATE)
        Employee.objects.filter(tab_number="0848103").update(method=DEG, voted=True)

        created, updated = import_base(TEMPLATE)
        self.assertEqual(created, 0)
        self.assertTrue(updated)

        person = Employee.objects.get(tab_number="0848103")
        self.assertEqual(person.method, DEG)
        self.assertTrue(person.voted)


class CountsTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_user("u", password="x")
        self.client.force_login(self.user)
        for i in range(4):
            Employee.objects.create(
                tab_number=f"t{i}",
                department="10" if i < 2 else "20",
                surname=f"Ф{i}",
                name="И",
                method=DEG if i == 0 else "",
            )

    def test_counts_follow_the_filter(self):
        page = self.client.get("/", {"dep": "10"})
        counts = page.context["counts"]["method"]
        self.assertEqual(counts["total"], 2)
        self.assertEqual(counts["deg"], 1)

        page = self.client.get("/", {"dep": "20"})
        counts = page.context["counts"]["method"]
        self.assertEqual(counts["total"], 2)
        self.assertEqual(counts["deg"], 0)

    def test_api_returns_counts_for_the_same_filter(self):
        target = Employee.objects.get(tab_number="t2")
        response = self.client.post(
            "/api/method/",
            {"id": target.pk, "method": DEG, "filters": {"dep": "20"}},
            content_type="application/json",
        )
        body = response.json()
        self.assertEqual(body["method"]["total"], 2)
        self.assertEqual(body["method"]["deg"], 1)
