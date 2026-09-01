"""
Модуль представлений (Views). Обрабатывает HTTP-запросы, управляет доступом и отдает ответы.
"""

import json
from typing import Any, Optional

from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q, QuerySet
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from .custom_reports import custom_production_summary, custom_report
from .helpers import COLUMNS
from .importers import import_base, import_voting_choices, mark_voted, set_turnout
from .models import DEG, METHOD_LABELS, METHODS, UIK, UIK19, UVZ, Employee
from .reports import (
    export_xlsx,
    production_method_table,
    production_table,
    reports_archive,
    summary_table,
    summary_table_no_u19,
)

# Константы
PER_PAGE = 100
CUSTOM_OKRUG_OPTIONS = [
    ("", "Все"),
    ("none", "Пусто"),
    ("19", "19"),
    ("20", "20"),
    ("21", "21"),
    ("20+21", "20+21"),
]
MARK_FIELDS = {"mark_deg": DEG, "mark_uvz": UVZ}


# ==============================================================================
# Вспомогательные функции (Helpers)
# ==============================================================================


def is_operator(user: Any) -> bool:
    """Проверяет, является ли пользователь оператором или суперюзером."""
    return user.is_superuser or user.groups.filter(name="operator").exists()


def can_edit(user: Any) -> bool:
    """Проверяет, может ли пользователь редактировать данные (не является viewer)."""
    return not user.groups.filter(name="viewer").exists()


def _known_method(value: Any) -> str:
    """Валидирует значение способа голосования через централизованный словарь METHOD_LABELS."""
    return str(value) if value in METHOD_LABELS else ""


def _parse_json_body(
    request: HttpRequest,
) -> tuple[Optional[dict], Optional[JsonResponse]]:
    """Безопасно парсит JSON из тела POST-запроса."""
    try:
        return json.loads(request.body or "{}"), None
    except json.JSONDecodeError:
        return None, JsonResponse({"error": "неверный формат запроса"}, status=400)


def _safe_int(value: Any) -> Optional[int]:
    """Безопасно преобразует значение в int. Возвращает None при ошибке."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _make_excel_response(workbook: Any, filename_prefix: str) -> HttpResponse:
    """DRY-хелпер для генерации HTTP-ответа с Excel-файлом."""
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    timestamp = timezone.localtime().strftime("%Y%m%d_%H%M")
    response["Content-Disposition"] = (
        f'attachment; filename="{filename_prefix}_{timestamp}.xlsx"'
    )
    workbook.save(response)
    return response


# ==============================================================================
# Фильтрация и статистика
# ==============================================================================


def _filtered(params: dict) -> QuerySet:
    """Строит QuerySet сотрудников на основе параметров фильтрации из GET-запроса."""
    qs = Employee.objects.all()
    search = (params.get("q") or "").strip()
    dep = (params.get("dep") or "").strip()
    production = (params.get("production") or "").strip()
    service = (params.get("service") or "").strip()
    okrug = (params.get("okrug") or "").strip()
    uik = (params.get("uik") or "").strip()
    method = (params.get("method") or "").strip()
    where = (params.get("where") or "").strip()
    voted = (params.get("voted") or "").strip()

    for part in search.split():
        qs = qs.filter(
            Q(tab_number__icontains=part)
            | Q(surname__icontains=part)
            | Q(name__icontains=part)
            | Q(patronymic__icontains=part)
        )

    if dep:
        qs = qs.filter(department=dep)
    if production:
        qs = qs.filter(production=production)
    if service:
        qs = qs.filter(service=service)
    if okrug == "none":
        qs = qs.filter(okrug="")
    elif okrug:
        qs = qs.filter(okrug=okrug)
    if uik:
        qs = qs.filter(uik=uik)
    if method == "none":
        qs = qs.filter(method="")
    elif method:
        qs = qs.filter(method=method)
    if where == "none":
        qs = qs.filter(voted_method="")
    elif where:
        qs = qs.filter(voted_method=where)
    if voted == "yes":
        qs = qs.filter(voted=True)
    elif voted == "no":
        qs = qs.filter(voted=False)

    return qs


def _counts(qs: Optional[QuerySet] = None) -> dict[str, Any]:
    """
    Считает агрегированную статистику.
    ОПТИМИЗАЦИЯ: Объединено в один запрос к БД вместо двух отдельных aggregate().
    """
    base = qs if qs is not None else Employee.objects.all()
    agg = base.aggregate(
        plan_deg=Count("id", filter=Q(method=DEG)),
        plan_uik=Count("id", filter=Q(method=UIK)),
        plan_uvz=Count("id", filter=Q(method=UVZ)),
        plan_u19=Count("id", filter=Q(method=UIK19)),
        plan_none=Count("id", filter=Q(method="")),
        plan_total=Count("id"),
        voted_deg=Count("id", filter=Q(voted=True, voted_method=DEG)),
        voted_uik=Count("id", filter=Q(voted=True, voted_method=UIK)),
        voted_uvz=Count("id", filter=Q(voted=True, voted_method=UVZ)),
        voted_u19=Count("id", filter=Q(voted=True, voted_method=UIK19)),
        voted_none=Count("id", filter=Q(voted=True, voted_method="")),
        voted_total=Count("id", filter=Q(voted=True)),
    )
    return {
        "method": {
            "deg": agg["plan_deg"],
            "uik": agg["plan_uik"],
            "uvz": agg["plan_uvz"],
            "u19": agg["plan_u19"],
            "none": agg["plan_none"],
            "total": agg["plan_total"],
        },
        "voted": {
            "deg": agg["voted_deg"],
            "uik": agg["voted_uik"],
            "uvz": agg["voted_uvz"],
            "u19": agg["voted_u19"],
            "none": agg["voted_none"],
            "total": agg["voted_total"],
        },
        "no_uik": base.filter(uik="").count(),
    }


def _page_window(page: Any, size: int = 5) -> range:
    """Вычисляет диапазон номеров страниц для пагинатора."""
    total = page.paginator.num_pages
    current = page.number
    if total <= size:
        return range(1, total + 1)
    half = size // 2
    start = max(1, current - half)
    end = min(total, current + half)
    if end - start + 1 < size:
        if start == 1:
            end = min(total, size)
        else:
            start = max(1, total - size + 1)
    return range(start, end + 1)


def _context(request: HttpRequest) -> dict[str, Any]:
    """Готовит общий контекст для страниц со списками сотрудников."""
    qs = _filtered(request.GET)
    page = Paginator(qs, PER_PAGE).get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)
    return {
        "page": page,
        "rows": page.object_list,
        "found": page.paginator.count,
        "counts": _counts(qs),
        "is_operator": is_operator(request.user),
        "methods": METHODS,
        "departments": Employee.objects.exclude(department="")
        .values_list("department", flat=True)
        .distinct()
        .order_by("department"),
        "productions": Employee.objects.exclude(production="")
        .values_list("production", flat=True)
        .distinct()
        .order_by("production"),
        "services": Employee.objects.exclude(service="")
        .values_list("service", flat=True)
        .distinct()
        .order_by("service"),
        "okrugs": Employee.objects.exclude(okrug="")
        .values_list("okrug", flat=True)
        .distinct()
        .order_by("okrug"),
        "uiks": Employee.objects.exclude(uik="")
        .values_list("uik", flat=True)
        .distinct()
        .order_by("uik"),
        "f": request.GET,
        "query": params.urlencode(),
        "page_range": _page_window(page),
        "can_edit": can_edit(request.user),
    }


# ==============================================================================
# Страницы (Views)
# ==============================================================================


@login_required
def method_page(request: HttpRequest) -> HttpResponse:
    """Страница со списком сотрудников для простановки способа голосования."""
    return render(request, "method.html", _context(request))


@login_required
def elections_page(request: HttpRequest) -> HttpResponse:
    """Страница со списком сотрудников для отметки явки."""
    return render(request, "elections.html", _context(request))


@login_required
def upload_page(request: HttpRequest) -> HttpResponse:
    """Страница загрузки файлов (доступна только операторам)."""
    if not is_operator(request.user):
        return render(request, "access_denied.html", {"is_operator": False})
    return render(
        request,
        "upload.html",
        {
            "counts": _counts(),
            "is_operator": True,
            "columns": list(COLUMNS),
            "msg": request.session.pop("msg", ""),
        },
    )


@login_required
def export_page(request: HttpRequest) -> HttpResponse:
    """Страница со списком доступных отчетов и формой конструктора."""
    return render(
        request,
        "export.html",
        {
            "counts": _counts(),
            "is_operator": is_operator(request.user),
            "departments_count": Employee.objects.exclude(department="")
            .values("department")
            .distinct()
            .count(),
            "productions_count": Employee.objects.exclude(production="")
            .values("service")
            .distinct()
            .count(),
            "msg": request.session.pop("msg", ""),
            "productions": Employee.objects.exclude(production="")
            .values_list("production", flat=True)
            .distinct()
            .order_by("production"),
            "services": Employee.objects.exclude(service="")
            .values_list("service", flat=True)
            .distinct()
            .order_by("service"),
            "departments": Employee.objects.exclude(department="")
            .values_list("department", flat=True)
            .distinct()
            .order_by("department"),
            "methods": METHODS,
            "uiks": Employee.objects.exclude(uik="")
            .values_list("uik", flat=True)
            .distinct()
            .order_by("uik"),
            "custom_okrugs": CUSTOM_OKRUG_OPTIONS,
        },
    )


# ==============================================================================
# Обработчики загрузки (Upload handlers)
# ==============================================================================


@login_required
@require_POST
def upload_base(request: HttpRequest) -> HttpResponse:
    """Обработчик загрузки основной базы сотрудников."""
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)
    upload = request.FILES.get("file")
    if not upload:
        return redirect("upload")

    # ИСПРАВЛЕНО: Валидация расширения на бэкенде
    if not upload.name.lower().endswith(".xlsx"):
        request.session["msg"] = "Ошибка: принимается только формат .xlsx"
        return redirect("upload")

    try:
        created, updated, total = import_base(upload)
        request.session["msg"] = (
            f"Строк в файле: {total}. Новых: {created}, изменено: {updated}, без изменений: {total - created - updated}"
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)
    return redirect("upload")


@login_required
@require_POST
def upload_voting_choices(request: HttpRequest) -> HttpResponse:
    """Обработчик загрузки отчета штаба для обновления способов голосования."""
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)
    upload = request.FILES.get("file")
    if not upload:
        return redirect("upload")

    # ИСПРАВЛЕНО: Валидация расширения на бэкенде
    if not upload.name.lower().endswith(".xlsx"):
        request.session["msg"] = "Ошибка: принимается только формат .xlsx"
        return redirect("upload")

    try:
        updated, total, errors = import_voting_choices(upload)
        request.session["msg"] = (
            f"Обработано строк: {total}. Обновлено способов: {updated}, ошибок/пропусков: {errors}"
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)
    return redirect("upload")


# ==============================================================================
# API endpoints
# ==============================================================================


@login_required
@require_POST
def api_method(request: HttpRequest) -> JsonResponse:
    """API: обновление способа голосования для одного сотрудника."""
    data, bad = _parse_json_body(request)
    if bad:
        return bad
    changed = Employee.objects.filter(pk=_safe_int(data.get("id"))).update(
        method=_known_method(data.get("method", ""))
    )
    if not changed:
        return JsonResponse({"error": "работник не найден"}, status=404)
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_mark(request: HttpRequest) -> JsonResponse:
    """API: простановка отметок (mark_deg, mark_uvz)."""
    data, bad = _parse_json_body(request)
    if bad:
        return bad
    field = data.get("field")
    if field not in MARK_FIELDS:
        return JsonResponse({"error": "неизвестное поле"}, status=400)
    person = Employee.objects.filter(pk=_safe_int(data.get("id"))).first()
    if person is None:
        return JsonResponse({"error": "работник не найден"}, status=404)
    if person.method != MARK_FIELDS[field]:
        return JsonResponse({"error": "способ не соответствует полю"}, status=400)
    setattr(person, field, bool(data.get("value")))
    person.save(update_fields=[field])
    return JsonResponse({field: getattr(person, field)})


@login_required
@require_POST
def api_voted(request: HttpRequest) -> JsonResponse:
    """API: отметка явки для одного сотрудника."""
    data, bad = _parse_json_body(request)
    if bad:
        return bad
    person = Employee.objects.filter(pk=_safe_int(data.get("id"))).first()
    if person is None:
        return JsonResponse({"error": "работник не найден"}, status=404)
    voted = bool(data.get("voted"))
    if voted and not person.method:
        return JsonResponse({"error": "Не выбран способ голосования"}, status=400)
    mark_voted([person.tab_number], voted=voted)
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_bulk_voted(request: HttpRequest) -> JsonResponse:
    """API: массовая отметка явки по текущим фильтрам."""
    data, bad = _parse_json_body(request)
    if bad:
        return bad
    voted = bool(data.get("voted"))
    filters = data.get("filters") or {}
    target = _filtered(filters)
    skipped = 0
    if voted:
        skipped = target.filter(method="").count()
        target = target.exclude(method="")
    changed = set_turnout(target, voted)
    result = _counts(_filtered(filters))
    result["changed"], result["skipped"] = changed, skipped
    return JsonResponse(result)


@login_required
def api_uik_stats(request: HttpRequest) -> JsonResponse:
    """API: статистика по УИКам для модального окна."""
    rows = (
        _filtered(request.GET)
        .values("uik")
        .annotate(people=Count("id"), came=Count("id", filter=Q(voted=True)))
        .order_by("uik")
    )
    return JsonResponse(
        list(rows), safe=False, json_dumps_params={"ensure_ascii": False}
    )


# ==============================================================================
# Экспорт отчетов (Standard)
# ==============================================================================


@login_required
def export_summary(request: HttpRequest) -> HttpResponse:
    return _make_excel_response(summary_table(), "svodka_po_ceham")


@login_required
def export_summary_no_19(request: HttpRequest) -> HttpResponse:
    return _make_excel_response(summary_table_no_u19(), "svodka_po_ceham_bez_u19")


@login_required
def export_productions(request: HttpRequest) -> HttpResponse:
    return _make_excel_response(production_table(), "po_proizvodstvam")


@login_required
def export_production_methods(request: HttpRequest) -> HttpResponse:
    return _make_excel_response(production_method_table(), "sposoby_po_proizvodstvam")


@login_required
def export_production_methods_no_19(request: HttpRequest) -> HttpResponse:
    return _make_excel_response(
        production_method_table(exclude_u19=True), "sposoby_po_proizvodstvam_bez_u19"
    )


@login_required
def export_employees(request: HttpRequest) -> HttpResponse:
    return _make_excel_response(export_xlsx(), "employees")


def _archive_response(request: HttpRequest, mode: str, prefix: str) -> HttpResponse:
    """Хелпер для генерации ZIP-архивов с отчетами по цехам."""
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
    return _archive_response(request, "turnout", "yavka")


@login_required
def export_method_archive(request: HttpRequest) -> HttpResponse:
    return _archive_response(request, "method", "sposob")


@login_required
def export_custom_report(request: HttpRequest) -> HttpResponse:
    """Экспорт кастомного отчета (по людям или по производствам)."""
    grouping = request.GET.get("grouping", "people")
    moment_str = timezone.localtime().strftime("%Y%m%d_%H%M")

    if grouping == "production_with_depts":
        book = custom_production_summary(request.GET, include_depts=True)
        name = f"svodny_po_proizvodstvam_s_cehami_{moment_str}"
    elif grouping == "production_without_depts":
        book = custom_production_summary(request.GET, include_depts=False)
        name = f"svodny_po_proizvodstvam_{moment_str}"
    else:
        book = custom_report(request.GET)
        name = f"svodny_otchet_{moment_str}"

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = f'attachment; filename="{name}.xlsx"'
    book.save(response)
    return response


# ==============================================================================
# Аутентификация
# ==============================================================================


def login_view(request: HttpRequest) -> HttpResponse:
    if request.user.is_authenticated:
        return redirect("method")
    error = False
    if request.method == "POST":
        user = authenticate(
            request,
            username=request.POST.get("username"),
            password=request.POST.get("password"),
        )
        if user:
            login(request, user)
            return redirect("method")
        error = True
    return render(request, "login.html", {"error": error})


@login_required
def logout_view(request: HttpRequest) -> HttpResponse:
    logout(request)
    return redirect("login")
