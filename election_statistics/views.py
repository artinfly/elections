import json

from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q
from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from .models import DEG, METHOD_LABELS, METHODS, UIK, UVZ, Employee
from .services import (
    COLUMNS,
    export_xlsx,
    import_base,
    mark_voted,
    production_table,
    reports_archive,
    set_turnout,
    summary_table,
)

PER_PAGE = 100


def is_operator(user):
    return user.is_superuser or user.groups.filter(name="operator").exists()


def _known_method(value):
    return value if value in METHOD_LABELS else ""


def _body(request):
    try:
        return json.loads(request.body or "{}"), None
    except ValueError:
        return None, JsonResponse({"error": "неверный формат запроса"}, status=400)


def _pk(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def login_view(request):
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


def logout_view(request):
    logout(request)
    return redirect("login")


def _filtered(params):
    qs = Employee.objects.all()
    search = (params.get("q") or "").strip()
    dep = (params.get("dep") or "").strip()
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


def _counts(qs=None):
    base = qs if qs is not None else Employee.objects.all()
    agg = base.aggregate(
        deg=Count("id", filter=Q(method=DEG)),
        uik=Count("id", filter=Q(method=UIK)),
        uvz=Count("id", filter=Q(method=UVZ)),
        none=Count("id", filter=Q(method="")),
        total=Count("id"),
    )
    voted = base.filter(voted=True).aggregate(
        deg=Count("id", filter=Q(voted_method=DEG)),
        uik=Count("id", filter=Q(voted_method=UIK)),
        uvz=Count("id", filter=Q(voted_method=UVZ)),
        none=Count("id", filter=Q(voted_method="")),
        total=Count("id"),
    )
    return {
        "method": agg,
        "voted": voted,
        "no_uik": base.filter(uik="").count(),
    }


def _context(request):
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
        "uiks": Employee.objects.exclude(uik="")
        .values_list("uik", flat=True)
        .distinct()
        .order_by("uik"),
        "f": request.GET,
        "query": params.urlencode(),
    }


@login_required
def method_page(request):
    return render(request, "method.html", _context(request))


@login_required
def elections_page(request):
    return render(request, "elections.html", _context(request))


@login_required
def upload_page(request):
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
@require_POST
def upload_base(request):
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав на загрузку файлов"}, status=403)
    upload = request.FILES.get("file")
    if not upload:
        return redirect("upload")
    try:
        created, updated, total = import_base(upload)
        request.session["msg"] = (
            f"Строк в файле: {total}. Новых работников: {created}, "
            f"изменено: {updated}, без изменений: {total - created - updated}"
        )
    except ValueError as exc:
        request.session["msg"] = str(exc)
    return redirect("upload")


@login_required
@require_POST
def api_method(request):
    data, bad = _body(request)
    if bad:
        return bad
    changed = Employee.objects.filter(pk=_pk(data.get("id"))).update(
        method=_known_method(data.get("method", ""))
    )
    if not changed:
        return JsonResponse({"error": "работник не найден"}, status=404)
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_voted(request):
    data, bad = _body(request)
    if bad:
        return bad
    person = Employee.objects.filter(pk=_pk(data.get("id"))).first()
    if person is None:
        return JsonResponse({"error": "работник не найден"}, status=404)

    voted = bool(data.get("voted"))
    if voted and not person.method:
        return JsonResponse(
            {"error": "У работника не выбран способ голосования"}, status=400
        )
    mark_voted([person.tab_number], voted=voted)
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_bulk_voted(request):
    data, bad = _body(request)
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
    result["changed"] = changed
    result["skipped"] = skipped
    return JsonResponse(result)


@login_required
def api_uik_stats(request):
    rows = (
        _filtered(request.GET)
        .values("uik")
        .annotate(people=Count("id"), came=Count("id", filter=Q(voted=True)))
        .order_by("uik")
    )
    return JsonResponse(
        list(rows), safe=False, json_dumps_params={"ensure_ascii": False}
    )


@login_required
def export_page(request):
    departments_count = (
        Employee.objects.exclude(department="").values("department").distinct().count()
    )
    productions_count = (
        Employee.objects.exclude(production="").values("production").distinct().count()
    )
    return render(
        request,
        "export.html",
        {
            "counts": _counts(),
            "is_operator": is_operator(request.user),
            "departments_count": departments_count,
            "productions_count": productions_count,
            "msg": request.session.pop("msg", ""),
        },
    )


@login_required
def export_summary(request):
    book = summary_table()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    name = f"svodka_po_ceham_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response


@login_required
def export_productions(request):
    book = production_table()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    name = f"po_proizvodstvam_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response


@login_required
def export_employees(request):
    book = export_xlsx()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = 'attachment; filename="employees.xlsx"'
    book.save(response)
    return response


def _archive_response(request, mode, prefix):
    moment = timezone.localtime()
    archiver = reports_archive(moment, mode)
    if not archiver.file_count:
        request.session["msg"] = "Нет ни одного отдела — архив пустой"
        return redirect("export")
    response = HttpResponse(archiver.build_bytes(), content_type="application/zip")
    name = f"{prefix}_po_ceham_{moment:%Y%m%d_%H%M}.zip"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    return response


@login_required
def export_archive(request):
    return _archive_response(request, "turnout", "yavka")


@login_required
def export_method_archive(request):
    return _archive_response(request, "method", "sposob")
