"""Импорт файлов: список сотрудников, явка штаба, кастомный отчёт, «через ответственного»."""

from datetime import datetime

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.utils.timezone import localtime

from core.models import DEG, UIK, UIK19, Employee
from core.services.importing import (
    BAD_ZIP,
    import_base,
    import_custom_report_archive,
    import_responsible_marks,
    import_turnout_hq_archive,
)

from .helpers import make_employee, xlsx, xlsx_bytes, zip_upload

BASE_HEADER = ["Подразделение", "Таб№", "Фамилия", "Имя", "Дата рождения", "Округ"]


class ImportBaseTests(TestCase):
    """Список сотрудников: создание, обновление, сохранность отметок."""

    def test_creates_updates_and_keeps_marks(self):
        make_employee(tab_number="0000001", surname="Старая", method=DEG, voted=True)
        upload = xlsx(
            [
                ["Выгрузка на 01.09"],  # шапка может быть не в первой строке
                BASE_HEADER,
                ["7", "1", "Новая", "Анна", "05.03.1990", "20"],  # 1 -> 0000001
                ["8", "123", "Петров", "Пётр", "", "21"],
            ]
        )
        self.assertEqual(import_base(upload), (1, 1, 2))
        updated = Employee.objects.get(tab_number="0000001")
        self.assertEqual(
            (updated.surname, updated.method, updated.voted), ("Новая", DEG, True)
        )
        self.assertTrue(Employee.objects.filter(tab_number="0000123").exists())

    def test_file_without_tab_column(self):
        with self.assertRaisesMessage(ValueError, "Таб№"):
            import_base(xlsx([["Подразделение", "Фамилия"], ["7", "Иванов"]]))


class TurnoutImportTests(TestCase):
    """Явка из архива штаба: время голосования и способ по умолчанию."""

    def test_marks_turnout_with_time(self):
        with_method = make_employee(method=DEG)
        without_method = make_employee(method="")
        book = xlsx_bytes(
            [
                ["Таб", "", "", "", "Дата", "Время"],
                [with_method.tab_number, "", "", "", "20.09.2026", "10:15"],
                [
                    without_method.tab_number,
                    "",
                    "",
                    "",
                    datetime(2026, 9, 19),
                    "09:00:00",
                ],
                ["не номер", "", "", "", "20.09.2026", "10:15"],
                ["9999999", "", "", "", "20.09.2026", "10:15"],
            ]
        )
        result = import_turnout_hq_archive(
            zip_upload({"1.xlsx": book, "__MACOSX/1.xlsx": b"x"})
        )
        self.assertEqual(result, (2, 4, 2))

        with_method.refresh_from_db()
        without_method.refresh_from_db()
        self.assertEqual(with_method.voted_method, DEG)
        self.assertEqual(
            (without_method.method, without_method.voted_method), (UIK, UIK)
        )
        self.assertEqual(
            localtime(with_method.voted_at).strftime("%d.%m %H:%M"), "20.09 10:15"
        )

    def test_bad_zip(self):
        with self.assertRaisesMessage(ValueError, BAD_ZIP):
            import_turnout_hq_archive(SimpleUploadedFile("x.zip", b"not zip"))


CUSTOM_GROUPS = [
    "Таб.№",
    "ДЭГ",
    "",
    "",
    "На участке",
    "",
    "УИК-19",
    "",
    "Уважительная причина",
]
CUSTOM_SUBS = [
    "",
    "Планирует",
    "Зарегистрирован",
    "Проголосовал",
    "Планирует",
    "Проголосовал",
    "Открепился",
    "Проголосовал",
    "",
]


class CustomReportImportTests(TestCase):
    """Кастомный отчёт: двухстрочная шапка, отметки, «Уважительная причина»."""

    def test_reads_marks(self):
        deg = make_employee()
        detached = make_employee()
        absent = make_employee()
        absent_but_voted = make_employee()
        upload = xlsx(
            [
                ["Сводный отчёт"],
                CUSTOM_GROUPS,
                CUSTOM_SUBS,
                [deg.tab_number, "", "+", "+", "", "", "", "", ""],
                [detached.tab_number, "", "", "", "", "", "+", "", ""],
                [absent.tab_number, "", "", "", "", "", "", "", "+"],
                [absent_but_voted.tab_number, "", "", "", "", "+", "", "", "+"],
                ["9999999", "+", "", "", "", "", "", "", ""],
            ]
        )
        self.assertEqual(import_custom_report_archive(upload), (4, 5, 1))

        deg.refresh_from_db()
        self.assertEqual((deg.method, deg.mark_deg, deg.voted_method), (DEG, True, DEG))
        detached.refresh_from_db()
        self.assertEqual((detached.method, detached.detached), (UIK19, True))
        absent.refresh_from_db()
        self.assertTrue(absent.absence)
        absent_but_voted.refresh_from_db()
        # Любая другая отметка важнее «Уважительной причины»
        self.assertFalse(absent_but_voted.absence)
        self.assertEqual(absent_but_voted.voted_method, UIK)

    def test_zip_with_several_books(self):
        first, second = make_employee(), make_employee()
        books = {
            f"{i}.xlsx": xlsx_bytes(
                [["x"], CUSTOM_GROUPS, CUSTOM_SUBS, [emp.tab_number, "+"]]
            )
            for i, emp in enumerate((first, second))
        }
        books["readme.txt"] = b"skip"
        self.assertEqual(import_custom_report_archive(zip_upload(books)), (2, 2, 0))


class ResponsibleImportTests(TestCase):
    """«Голосование через ответственного»: заголовки-синонимы и время отметки."""

    def test_marks_and_keeps_voting_time(self):
        voted_at = datetime(2026, 9, 19, 9, 0).astimezone()
        voted = make_employee(voted=True, voted_at=voted_at)
        not_voted = make_employee()
        skipped = make_employee()
        upload = xlsx(
            [
                ["Цех", "Таб №", "ФИО", "отм."],
                ["7", voted.tab_number, "", "1"],
                ["7", not_voted.tab_number, "", 0],  # 0 — тоже отметка
                ["7", skipped.tab_number, "", ""],
            ]
        )
        self.assertEqual(import_responsible_marks(upload), (2, 3, 0))

        voted.refresh_from_db()
        not_voted.refresh_from_db()
        skipped.refresh_from_db()
        self.assertTrue(voted.proxy_vote and not_voted.proxy_vote)
        self.assertFalse(skipped.proxy_vote)
        self.assertEqual(voted.voted_at, voted_at)
        self.assertIsNotNone(not_voted.voted_at)
