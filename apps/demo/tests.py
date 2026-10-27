from io import StringIO

from django.core.management import call_command
from django.test import TestCase

from core.models import Employee


class SeedDemoTests(TestCase):
    def test_seed(self):
        call_command("seed_demo", "--people", "80", stdout=StringIO())
        self.assertEqual(Employee.objects.count(), 80)
        self.assertTrue(Employee.objects.exclude(method="").exists())
        self.assertTrue(Employee.objects.filter(voted=True).exists())
        self.assertFalse(Employee.objects.filter(production="").exists())
        # Повторный запуск заменяет данные, а не добавляет
        call_command("seed_demo", "--people", "50", stdout=StringIO())
        self.assertEqual(Employee.objects.count(), 50)
