"""API кнопок: способ, регистрация, явка, отсутствие по УП."""

import json

from django.test import TestCase
from django.urls import reverse

from core.models import DEG, UIK, UVZ

from .helpers import make_employee, make_user


class ApiTests(TestCase):
    """Отметки без перезагрузки страницы."""

    def setUp(self):
        self.client.force_login(make_user("user"))

    def post(self, name, data, **kwargs):
        return self.client.post(
            reverse(name, kwargs=kwargs),
            json.dumps(data),
            content_type="application/json",
        )

    def test_set_method_returns_counts(self):
        employee = make_employee()
        response = self.post("api_method", {"id": employee.id, "method": DEG})
        employee.refresh_from_db()
        self.assertEqual(employee.method, DEG)
        self.assertEqual(response.json()["method"]["deg"], 1)

    def test_unknown_method_is_cleared(self):
        employee = make_employee(method=DEG)
        self.post("api_method", {"id": employee.id, "method": "xxx"})
        employee.refresh_from_db()
        self.assertEqual(employee.method, "")

    def test_bad_requests(self):
        self.assertEqual(self.post("api_method", {"id": "abc"}).status_code, 400)
        self.assertEqual(
            self.post("api_method", {"id": 99999, "method": DEG}).status_code, 404
        )
        response = self.client.post(
            reverse("api_voted"), "{", content_type="application/json"
        )
        self.assertEqual(response.status_code, 400)

    def test_registration_mark_needs_matching_method(self):
        employee = make_employee(method=UIK)
        self.assertEqual(
            self.post(
                "api_mark", {"id": employee.id, "field": "mark_deg", "value": True}
            ).status_code,
            400,
        )
        employee.method = UVZ
        employee.save()
        response = self.post(
            "api_mark", {"id": employee.id, "field": "mark_uvz", "value": True}
        )
        self.assertEqual(response.json(), {"mark_uvz": True})

    def test_voted_copies_method_and_unvote_clears_it(self):
        employee = make_employee(method=DEG)
        self.post("api_voted", {"id": employee.id, "voted": True})
        employee.refresh_from_db()
        self.assertTrue(employee.voted)
        self.assertEqual(employee.voted_method, DEG)
        self.assertIsNotNone(employee.voted_at)

        self.post("api_voted", {"id": employee.id, "voted": False})
        employee.refresh_from_db()
        self.assertFalse(employee.voted)
        self.assertEqual(employee.voted_method, "")

    def test_voted_without_method_sets_uik(self):
        employee = make_employee(method="")
        self.post("api_voted", {"id": employee.id, "voted": True})
        employee.refresh_from_db()
        self.assertEqual((employee.method, employee.voted_method), (UIK, UIK))

    def test_bulk_voted_skips_people_without_method(self):
        make_employee(department="5", method=DEG)
        make_employee(department="5", method="")
        make_employee(department="6", method=DEG)
        response = self.post("api_bulk_voted", {"voted": True, "filters": {"dep": "5"}})
        self.assertEqual(
            (response.json()["changed"], response.json()["skipped"]), (1, 1)
        )

    def test_toggle_absence(self):
        employee = make_employee()
        response = self.post("toggle_absence", {}, employee_id=employee.id)
        self.assertEqual(response.json(), {"id": employee.id, "absence": True})

    def test_uik_stats(self):
        make_employee(uik="1901", voted=True)
        make_employee(uik="1901")
        response = self.client.get(reverse("api_uik_stats"))
        self.assertEqual(response.json(), [{"uik": "1901", "people": 2, "came": 1}])
