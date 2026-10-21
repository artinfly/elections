"""HTML-страницы: способ голосования, явка, загрузка, отчёты."""

from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Q
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect, render

from ..models import METHODS, UVZ, Employee
from ..services.excel import COLUMNS
from .common import (
    PER_PAGE,
    filter_employees,
    filter_options,
    is_operator,
    is_uik_uvz_viewer,
    list_context,
    vote_counts,
)

# Варианты фильтра по избирательным округам в конструкторе отчётов.
CUSTOM_OKRUG_OPTIONS = [
    ("", "Все"),
    ("none", "Пусто"),
    ("19", "19"),
    ("20", "20"),
    ("21", "21"),
    ("20+21", "20+21"),
]


@login_required
def method_page(request: HttpRequest) -> HttpResponse:
    """Главная страница: список сотрудников для простановки способа голосования."""
    if is_uik_uvz_viewer(request.user):
        return redirect("elections")
    return render(request, "pages/method.html", list_context(request))


@login_required
def elections_page(request: HttpRequest) -> HttpResponse:
    """Страница выборов: список сотрудников для отметки фактической явки."""
    context = list_context(request)

    if is_uik_uvz_viewer(request.user):
        qs = filter_employees(request.GET).filter(Q(method=UVZ) | Q(voted_method=UVZ))
        page = Paginator(qs, PER_PAGE).get_page(request.GET.get("page"))

        context["page"] = page
        context["rows"] = page.object_list
        context["found"] = page.paginator.count
        context["counts"] = vote_counts(qs)

    return render(request, "pages/elections.html", context)


@login_required
def upload_page(request: HttpRequest) -> HttpResponse:
    """Страница загрузки файлов из Excel."""
    if is_uik_uvz_viewer(request.user):
        return redirect("elections")
    return render(
        request,
        "pages/upload.html",
        {
            "counts": vote_counts(),
            "is_operator": is_operator(request.user),
            "columns": list(COLUMNS),
            "msg": request.session.pop("msg", ""),
        },
    )


@login_required
def export_page(request: HttpRequest) -> HttpResponse:
    """Страница экспорта: список готовых отчетов и конструктор кастомных сводок."""
    if is_uik_uvz_viewer(request.user):
        return redirect("elections")
    options = filter_options()

    return render(
        request,
        "pages/export.html",
        {
            "counts": vote_counts(),
            "is_operator": is_operator(request.user),
            "departments_count": len(options["departments"]),
            "productions_count": len(options["productions"]),
            "msg": request.session.pop("msg", ""),
            "productions": options["productions"],
            "services": options["services"],
            "departments": options["departments"],
            "methods": METHODS,
            "uiks": options["uiks"],
            "custom_okrugs": CUSTOM_OKRUG_OPTIONS,
            # Сколько человек оформили заявление на УВЗ, но не проголосовали.
            "uvz_not_voted_count": Employee.objects.filter(
                method=UVZ, mark_uvz=True, voted=False
            ).count(),
        },
    )
