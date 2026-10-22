"""Общие помощники тестов: создание сотрудников, пользователей и Excel-файлов."""

import io
import zipfile
from itertools import count

import openpyxl
from django.contrib.auth.models import Group, User
from django.core.files.uploadedfile import SimpleUploadedFile

from core.models import Employee

_numbers = count(1)


def make_employee(**fields) -> Employee:
    """Сотрудник с уникальным семизначным табельным номером и заданными полями."""
    number = next(_numbers)
    defaults = {
        "tab_number": f"{1000000 + number:07d}",
        "department": "7",
        "surname": "Иванов",
        "name": f"Иван{number}",
        "patronymic": "Иванович",
    }
    defaults.update(fields)
    return Employee.objects.create(**defaults)


def make_user(username: str, *groups: str, superuser: bool = False) -> User:
    """Пользователь с паролем «pass» в указанных группах ролей."""
    if superuser:
        user = User.objects.create_superuser(username, password="pass")
    else:
        user = User.objects.create_user(username, password="pass")
    for name in groups:
        user.groups.add(Group.objects.get_or_create(name=name)[0])
    return user


def xlsx(rows: list, name: str = "file.xlsx") -> SimpleUploadedFile:
    """Загружаемый .xlsx с одним листом из переданных строк."""
    return SimpleUploadedFile(name, xlsx_bytes(rows))


def xlsx_bytes(rows: list) -> bytes:
    """Содержимое .xlsx с одним листом из переданных строк."""
    book = openpyxl.Workbook()
    for row in rows:
        book.active.append(row)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def zip_upload(files: dict, name: str = "archive.zip") -> SimpleUploadedFile:
    """Загружаемый .zip из словаря {имя файла: байты}."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for filename, data in files.items():
            archive.writestr(filename, data)
    return SimpleUploadedFile(name, buffer.getvalue())


def read_book(content: bytes) -> openpyxl.Workbook:
    """Книга Excel из ответа сервера."""
    return openpyxl.load_workbook(io.BytesIO(content))
