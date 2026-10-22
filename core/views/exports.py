"""Выгрузки: Excel-отчёты и ZIP-архивы по цехам."""

from typing import Any

from django.contrib.auth.decorators import login_required
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect
from django.utils import timezone

from ..services.reports.custom import custom_production_summary, custom_report
from ..services.reports.standard import (
    custom_reports_archive,
    export_xlsx,
    production_method_table,
    production_table,
    reports_archive,
    responsible_marks_archive,
    summary_table,
)
from ..services.reports.uvz import uvz_statement_not_voted_table
from ..services.reports.vn import custom_production_summary_vn
from .common import is_operator


def _make_excel_response(workbook: Any, filename_prefix: str) -> HttpResponse:
    """
    Формирует HTTP-ответ для скачивания Excel-файла браузером.

    Аргументы:
        workbook: объект книги openpyxl.
        filename_prefix: базовая часть имени файла (например, "svodka").

    Возвращает:
        HttpResponse с прикрепленным Excel-файлом.
    """
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    timestamp = timezone.localtime().strftime("%Y%m%d_%H%M")
    response["Content-Disposition"] = (
        f'attachment; filename="{filename_prefix}_{timestamp}.xlsx"'
    )
    workbook.save(response)
    return response


@login_required
def export_summary(request: HttpRequest) -> HttpResponse:
    """Экспорт сводной таблицы по цехам."""
    return _make_excel_response(summary_table(), "svodka_po_ceham")


@login_required
def export_summary_no_19(request: HttpRequest) -> HttpResponse:
    """Экспорт сводной таблицы по цехам без учета 19 округа."""
    return _make_excel_response(
        summary_table(exclude_u19=True), "svodka_po_ceham_bez_u19"
    )


@login_required
def export_productions(request: HttpRequest) -> HttpResponse:
    """Экспорт таблицы по производствам."""
    return _make_excel_response(production_table(), "po_proizvodstvam")


@login_required
def export_production_methods(request: HttpRequest) -> HttpResponse:
    """Экспорт таблицы способов голосования по производствам."""
    return _make_excel_response(production_method_table(), "sposoby_po_proizvodstvam")


@login_required
def export_production_methods_no_19(request: HttpRequest) -> HttpResponse:
    """Экспорт таблицы способов голосования по производствам без 19 округа."""
    return _make_excel_response(
        production_method_table(exclude_u19=True), "sposoby_po_proizvodstvam_bez_u19"
    )


@login_required
def export_employees(request: HttpRequest) -> HttpResponse:
    """Экспорт полной таблицы сотрудников."""
    return _make_excel_response(export_xlsx(), "employees")


@login_required
def export_uvz_statement_not_voted(request: HttpRequest) -> HttpResponse:
    """Экспорт списка: заявление на УИК-УВЗ оформлено, но не проголосовал."""
    return _make_excel_response(
        uvz_statement_not_voted_table(), "spisok_ne_progolosovavshih_na_UIK-UVZ"
    )


@login_required
def export_responsible_template(request: HttpRequest) -> HttpResponse:
    """Экспорт шаблона для заполнения отметок «Голосование через ответственного»."""
    if not is_operator(request.user):
        request.session["msg"] = "Нет прав для выполнения этого действия"
        return redirect("export")

    moment = timezone.localtime()
    archiver = responsible_marks_archive()

    if not archiver.file_count:
        request.session["msg"] = "Нет данных для формирования архива"
        return redirect("export")

    response = HttpResponse(archiver.build_bytes(), content_type="application/zip")
    response["Content-Disposition"] = (
        f'attachment; filename="shablon_otvetstvenny_{moment:%Y%m%d_%H%M}.zip"'
    )
    return response


def _archive_response(request: HttpRequest, mode: str, prefix: str) -> HttpResponse:
    """Внутренний хелпер для генерации ZIP-архивов с отчетами по цехам."""
    moment = timezone.localtime()
    archiver = reports_archive(moment, mode)

    if not archiver.file_count:
        request.session["msg"] = "Нет ни одного отдела — архив пустой"
        return redirect("export")

    response = HttpResponse(archiver.build_bytes(), content_type="application/zip")
    response["Content-Disposition"] = (
        f'attachment; filename="{prefix}_po_ceham_{moment:%Y%m%d_%H%M}.zip"'
    )
    return response


@login_required
def export_archive(request: HttpRequest) -> HttpResponse:
    """Экспорт архива отчетов по явке."""
    return _archive_response(request, "turnout", "yavka")


@login_required
def export_custom_archive(request: HttpRequest) -> HttpResponse:
    """Экспорт кастомного архива отчетов по цехам."""
    moment = timezone.localtime()
    archiver = custom_reports_archive(request.GET)

    if not archiver.file_count:
        request.session["msg"] = "Нет ни одного цеха - архив пустой"
        return redirect("export")

    response = HttpResponse(archiver.build_bytes(), content_type="application/zip")
    response["Content-Disposition"] = (
        f'attachment; filename="svodny_po_ceham_{moment:%Y%m%d_%H%M}.zip"'
    )
    return response


@login_required
def export_method_archive(request: HttpRequest) -> HttpResponse:
    """Экспорт архива отчетов по способам голосования."""
    return _archive_response(request, "method", "sposob")


@login_required
def export_custom_report(request: HttpRequest) -> HttpResponse:
    """Экспорт кастомного отчета (по людям или по производствам, включая ВН)."""
    grouping = request.GET.get("grouping", "people")
    moment_str = timezone.localtime().strftime("%Y%m%d_%H%M")

    if grouping == "production_with_depts":
        book = custom_production_summary(request.GET, include_depts=True)
        name = f"svodny_po_proizvodstvam_s_cehami_{moment_str}"
    elif grouping == "production_without_depts":
        book = custom_production_summary(request.GET, include_depts=False)
        name = f"svodny_po_proizvodstvam_{moment_str}"
    elif grouping == "people_compact":
        book = custom_report(request.GET, short=True)
        name = f"svodny_otchet_{moment_str}"
    elif grouping == "production_with_depts_compact":
        book = custom_production_summary(request.GET, include_depts=True, short=True)
        name = f"svodny_po_proizvodstvam_s_cehami_{moment_str}"
    elif grouping == "production_without_depts_compact":
        book = custom_production_summary(request.GET, include_depts=False, short=True)
        name = f"svodny_po_proizvodstvam_{moment_str}"
    elif grouping == "production_with_depts_vn":
        book = custom_production_summary_vn(request.GET, include_depts=True)
        name = f"svodny_po_proizvodstvam_s_cehami_vn_{moment_str}"
    elif grouping == "production_without_depts_vn":
        book = custom_production_summary_vn(request.GET, include_depts=False)
        name = f"svodny_po_proizvodstvam_vn_{moment_str}"
    else:
        book = custom_report(request.GET)
        name = f"svodny_otchet_{moment_str}"

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = f'attachment; filename="{name}.xlsx"'
    book.save(response)

    return response
