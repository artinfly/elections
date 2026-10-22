"""Модели и утилиты чтения/оформления Excel."""

from datetime import date, datetime

from django.test import SimpleTestCase, TestCase

from core.admin import archive_fields
from core.models import DEG, EmployeeArchive
from core.services.excel import (
    _by_number,
    _date,
    _format_with_percent,
    _text,
    department_file_name,
    padded_number,
)

from .helpers import make_employee


class EmployeeTests(TestCase):
    """Свойства сотрудника и перенос в архив."""

    def test_fio_skips_empty_parts(self):
        employee = make_employee(surname="Петров", name="Пётр", patronymic="")
        self.assertEqual(employee.fio, "Петров Пётр")

    def test_method_labels(self):
        employee = make_employee(method=DEG, voted_method="")
        self.assertEqual(employee.method_label, "ДЭГ")
        self.assertEqual(employee.voted_method_label, "")

    def test_archive_copies_every_field(self):
        employee = make_employee(uik="1901", id_uik=55, proxy_vote=True, okrug="20")
        archived = EmployeeArchive.objects.create(**archive_fields(employee))
        for field in ("tab_number", "uik", "id_uik", "proxy_vote", "okrug", "surname"):
            self.assertEqual(getattr(archived, field), getattr(employee, field), field)


class ExcelUtilsTests(SimpleTestCase):
    """Приведение значений ячеек и оформление номеров."""

    def test_text_turns_whole_floats_into_ints(self):
        self.assertEqual(_text(1234567.0), "1234567")
        self.assertEqual(_text("  abc "), "abc")
        self.assertEqual(_text(None), "")

    def test_date_formats(self):
        self.assertEqual(_date("05.03.1990"), date(1990, 3, 5))
        self.assertEqual(_date("1990-03-05 00:00:00"), date(1990, 3, 5))
        self.assertEqual(_date(datetime(1990, 3, 5, 12)), date(1990, 3, 5))
        self.assertIsNone(_date("вчера"))

    def test_department_numbers(self):
        self.assertEqual(padded_number("7"), "007")
        self.assertEqual(padded_number("АУП"), "АУП")
        self.assertEqual(department_file_name("1/2"), "1-2")
        self.assertEqual(
            sorted(["10", "2", "Б", "А"], key=_by_number), ["2", "10", "А", "Б"]
        )

    def test_percent_format(self):
        self.assertEqual(_format_with_percent(1, 4), "1 (25%)")
        self.assertEqual(_format_with_percent(3, 0), "3")
