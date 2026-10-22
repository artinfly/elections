"""Отчёты Excel и архивы: отдаются, итоги сходятся с данными."""

import io
import zipfile

from django.test import TestCase
from django.urls import reverse

from core.models import DEG, UIK, UIK19, UVZ
from core.services.reports.custom import _okrug19_total, _okrug19_votes

from .helpers import make_employee, make_user, read_book

EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class ReportTestCase(TestCase):
    """Небольшой набор сотрудников в двух цехах и двух производствах."""

    @classmethod
    def setUpTestData(cls):
        make_employee(
            department="7",
            service="Литейное",
            method=DEG,
            voted=True,
            voted_method=DEG,
            okrug="20",
        )
        make_employee(department="7", service="Литейное", method=UIK, okrug="21")
        make_employee(
            department="12",
            service="Сварочное",
            method=UVZ,
            voted=True,
            voted_method=UVZ,
            okrug="19",
        )
        make_employee(
            department="12", service="Сварочное", method="", absence=True, okrug=""
        )
        cls.user = make_user("user")

    def setUp(self):
        self.client.force_login(self.user)


class ExportTests(ReportTestCase):
    """Каждая выгрузка отдаёт файл нужного типа."""

    def test_all_exports_respond(self):
        excel = [
            "export_employees",
            "export_summary",
            "export_summary_no_19",
            "export_productions",
            "export_production_methods",
            "export_production_methods_no_19",
            "export_uvz_statement_not_voted",
            "export_custom",
        ]
        for name in excel:
            response = self.client.get(reverse(name))
            self.assertEqual(response["Content-Type"], EXCEL, name)

        for name in (
            "export_archive",
            "export_method_archive",
            "export_custom_archive",
        ):
            response = self.client.get(reverse(name))
            with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
                self.assertEqual(
                    sorted(archive.namelist()), ["007.xlsx", "012.xlsx"], name
                )

    def test_custom_report_groupings(self):
        groupings = (
            "people",
            "people_compact",
            "production_with_depts",
            "production_without_depts",
            "production_with_depts_compact",
            "production_without_depts_compact",
            "production_with_depts_vn",
            "production_without_depts_vn",
        )
        for grouping in groupings:
            response = self.client.get(reverse("export_custom"), {"grouping": grouping})
            self.assertEqual(response.status_code, 200, grouping)

    def test_responsible_template_needs_operator(self):
        response = self.client.get(reverse("export_responsible_template"))
        self.assertRedirects(response, reverse("export"))
        self.client.force_login(make_user("op", "operator"))
        self.assertEqual(
            self.client.get(reverse("export_responsible_template")).status_code, 200
        )


class SummaryTests(ReportTestCase):
    """Сводная таблица по цехам: строки и «Итого»."""

    def test_rows_and_totals(self):
        sheet = read_book(self.client.get(reverse("export_summary")).content).active
        rows = [[c.value for c in row] for row in sheet.iter_rows(min_row=2)]
        self.assertEqual(rows[0][:8], ["007", 2, 1, 1, 0, 0, 1, 0])
        self.assertEqual(rows[1][:8], ["012", 2, 0, 0, 1, 0, 1, 0])
        self.assertEqual(rows[2][:8], ["Итого", 4, 1, 1, 1, 0, 2, 0])
        self.assertEqual(rows[2][8], 0.5)

    def test_without_okrug_19(self):
        sheet = read_book(
            self.client.get(reverse("export_summary_no_19")).content
        ).active
        total = [c.value for c in list(sheet.iter_rows())[-1]]
        self.assertEqual(total[:2], ["Итого", 3])


class ProductionReportTests(ReportTestCase):
    """Отчёты «по производствам» и колонки 19 округа."""

    def test_productions_blocks(self):
        sheet = read_book(self.client.get(reverse("export_productions")).content).active
        values = [
            c.value for row in sheet.iter_rows() for c in row if c.value is not None
        ]
        self.assertIn("Литейное", values)
        self.assertIn("Сварочное", values)
        last = [c.value for c in list(sheet.iter_rows())[-1]]
        self.assertEqual(last[:4], ["Всего", 4, 2, 0])

    def test_okrug19_rule(self):
        # «не 19 округ»: округ 20/21, либо пустой округ не в исключённых городах
        person = make_employee(okrug="20", method=DEG)
        self.assertFalse(_okrug19_total(person))
        person = make_employee(okrug="19", method=UVZ, voted=True, voted_method=UVZ)
        self.assertTrue(_okrug19_total(person))
        self.assertTrue(_okrug19_votes(person))
        person = make_employee(okrug="", city="Нижний Тагил", method=UIK19)
        self.assertTrue(_okrug19_total(person))

    def test_vn_excludes_absent_who_did_not_vote(self):
        response = self.client.get(
            reverse("export_custom"), {"grouping": "production_without_depts_vn"}
        )
        sheet = read_book(response.content).active
        last = [c.value for c in list(sheet.iter_rows())[-1]]
        self.assertEqual(last[:2], ["Всего по Обществу", 3])
