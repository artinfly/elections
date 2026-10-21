"""Приём файлов со страницы «Загрузка»: только для оператора и суперпользователя."""

from typing import Any

from django.contrib.auth.decorators import login_required
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect
from django.views.decorators.http import require_POST

from ..services.importing import (
    import_base,
    import_custom_report_archive,
    import_responsible_marks,
    import_turnout_hq_archive,
)
from .common import is_operator


def _deny_upload(request: HttpRequest) -> HttpResponse:
    """Загрузка без прав оператора: сообщение на странице загрузки вместо ошибки."""
    request.session["msg"] = "Ошибка: загружать файлы может только оператор"
    return redirect("upload")


def _validate_excel_file(file: Any) -> None:
    """
    Проверяет расширение загружаемого файла.

    Аргументы:
        file: объект загруженного файла из request.FILES.

    Исключения:
        ValueError: если файл отсутствует или имеет неверное расширение.
    """
    if not file or not file.name.lower().endswith(".xlsx"):
        raise ValueError("Ошибка: принимается только формат .xlsx")


@login_required
@require_POST
def upload_base(request: HttpRequest) -> HttpResponse:
    """Обработчик формы загрузки основного списка сотрудников."""
    if not is_operator(request.user):
        return _deny_upload(request)

    upload = request.FILES.get("file")
    try:
        _validate_excel_file(upload)
    except ValueError as exc:
        request.session["msg"] = str(exc)
        return redirect("upload")

    try:
        created, updated, total = import_base(upload)
        request.session["msg"] = (
            f"Строк в файле: {total}. Новых: {created}, изменено: {updated}, "
            f"без изменений: {total - created - updated}"
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)
    return redirect("upload")


@login_required
@require_POST
def upload_turnout_hq(request: HttpRequest) -> HttpResponse:
    """Обработчик формы загрузки отметок явки из файла штаба (с датой и временем)."""
    if not is_operator(request.user):
        return _deny_upload(request)

    upload = request.FILES.get("file")
    if not upload or not upload.name.lower().endswith(".zip"):
        request.session["msg"] = "Ошибка: принимается только формат .zip"
        return redirect("upload")

    try:
        changed, total, errors = import_turnout_hq_archive(upload)
        request.session["msg"] = (
            f"Обработано файлов в архиве. Всего строк {total}. "
            f"Успешно отмечено явок: {changed}, ошибок/пропусков: {errors}"
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)
    except Exception as exc:
        request.session["msg"] = f"Неожиданная ошибка: {str(exc)}"

    return redirect("upload")


@login_required
@require_POST
def upload_custom_report(request: HttpRequest) -> HttpResponse:
    """Обработчик формы загрузки кастомного сводного отчёта."""
    if not is_operator(request.user):
        return _deny_upload(request)

    upload = request.FILES.get("file")
    if not upload or not (
        upload.name.lower().endswith(".zip") or upload.name.lower().endswith(".xlsx")
    ):
        request.session["msg"] = "Ошибка: принимается только формат .zip или .xlsx"
        return redirect("upload")

    try:
        changed, total, errors = import_custom_report_archive(upload)
        request.session["msg"] = (
            f"Обработано строк: {total}. Успешно обновлено записей: {changed}. "
            f"Ошибок/пропусков (нет в базе): {errors}."
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)
    except Exception as exc:
        request.session["msg"] = f"Неожиданная ошибка: {str(exc)}"

    return redirect("upload")


@login_required
@require_POST
def upload_responsible_marks(request: HttpRequest) -> HttpResponse:
    """Обработчик формы загрузки отметок «Голосование через ответственного»."""
    if not is_operator(request.user):
        return _deny_upload(request)

    upload = request.FILES.get("file")
    if not upload or not (
        upload.name.lower().endswith(".zip") or upload.name.lower().endswith(".xlsx")
    ):
        request.session["msg"] = "Ошибка: принимается только формат .zip или .xlsx"
        return redirect("upload")

    try:
        changed, total, errors = import_responsible_marks(upload)
        request.session["msg"] = (
            f"Обработано строк: {total}. Отмечено голосований через "
            f"ответственного: {changed}. Ошибок/пропусков: {errors}."
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)

    return redirect("upload")
