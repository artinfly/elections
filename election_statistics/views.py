import json

from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q
from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from .models import DEG, METHODS, UIK, UVZ, Employee
from .services import (
    export_xlsx,
    import_base,
    import_methods,
    mark_voted,
    parse_tabs,
    read_tab_column,
)

PER_PAGE = 100


def is_operator(user):
    return user.is_superuser or user.groups.filter(name="operator").exists()


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
    tab = (params.get("tab") or "").strip()
    fio = (params.get("fio") or "").strip()
    dep = (params.get("dep") or "").strip()
    uik = (params.get("uik") or "").strip()
    method = (params.get("method") or "").strip()
    voted = (params.get("voted") or "").strip()

    if tab:
        qs = qs.filter(tab_number__icontains=tab)
    if fio:
        for part in fio.split():
            qs = qs.filter(
                Q(surname__icontains=part)
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
        deg=Count("id", filter=Q(method=DEG)),
        uik=Count("id", filter=Q(method=UIK)),
        uvz=Count("id", filter=Q(method=UVZ)),
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
    return {
        "page": page,
        "rows": page.object_list,
        "found": qs.count(),
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
        "query": request.GET.urlencode(),
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
        created, updated = import_base(upload)
        request.session["msg"] = f"Загружено: новых {created}, обновлено {updated}"
    except ValueError as exc:
        request.session["msg"] = str(exc)
    return redirect("upload")


@login_required
@require_POST
def upload_methods(request):
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав на загрузку файлов"}, status=403)
    upload = request.FILES.get("file")
    if not upload:
        return redirect("upload")
    try:
        changed, skipped = import_methods(upload)
        request.session["msg"] = f"Способ проставлен: {changed}, пропущено {skipped}"
    except ValueError as exc:
        request.session["msg"] = str(exc)
    return redirect("upload")


@login_required
@require_POST
def upload_voted(request):
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав на загрузку файлов"}, status=403)
    tabs = []
    upload = request.FILES.get("file")
    if upload:
        try:
            tabs = read_tab_column(upload)
        except ValueError as exc:
            request.session["msg"] = str(exc)
            return redirect("upload")
    tabs += parse_tabs(request.POST.get("tabs", ""))
    changed, missing = mark_voted(tabs)
    request.session["msg"] = f"Отмечено: {changed}, не найдено {missing}"
    return redirect("upload")


@login_required
@require_POST
def api_method(request):
    data = json.loads(request.body or "{}")
    Employee.objects.filter(pk=data.get("id")).update(method=data.get("method", ""))
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_voted(request):
    data = json.loads(request.body or "{}")
    person = Employee.objects.filter(pk=data.get("id")).first()
    if person:
        mark_voted([person.tab_number], voted=bool(data.get("voted")))
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_bulk_voted(request):
    data = json.loads(request.body or "{}")
    voted = bool(data.get("voted"))
    filters = data.get("filters") or {}
    tabs = list(_filtered(filters).values_list("tab_number", flat=True))
    changed, _ = mark_voted(tabs, voted=voted)
    result = _counts(_filtered(filters))
    result["changed"] = changed
    return JsonResponse(result)


@login_required
def export_page(request):
    book = export_xlsx()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = 'attachment; filename="employees.xlsx"'
    book.save(response)
    return response
