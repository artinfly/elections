import json

from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q
from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from .models import DEG, METHOD_LABELS, METHODS, UIK, UIK19, UVZ, Employee
from .services import (
    COLUMNS,
    custom_production_summary,
    custom_report,
    export_xlsx,
    import_base,
    mark_voted,
    production_method_table,
    production_table,
    reports_archive,
    set_turnout,
    summary_table,
    summary_table_no_u19,
)

# Количество сотрудников на одной странице пагинации
PER_PAGE = 100

# Опции для селекта "Округ" в форме конструктора сводного отчёта
CUSTOM_OKRUG_OPTIONS = [
    ("", "Все"),
    ("none", "Пусто"),
    ("19", "19"),
    ("20", "20"),
    ("21", "21"),
    ("20+21", "20+21"),
]


def is_operator(user):
    """
    Проверяет, является ли пользователь оператором.
    Оператор — это суперюзер или пользователь, входящий в группу 'operator'.
    Операторы имеют доступ к загрузке базы и другим расширенным функциям.
    """
    return user.is_superuser or user.groups.filter(name="operator").exists()


def can_edit(user):
    """
    Может ли пользователь редактировать способы голосования и явку.
    Вьюеры (группа 'viewer') - только просматривают, но не меняют.
    """
    return not user.groups.filter(name="viewer").exists()


def _known_method(value):
    """
    Валидирует значение способа голосования.
    Возвращает код способа, если он есть в справочнике METHOD_LABELS, иначе пустую строку.
    Защищает от записи невалидных данных с фронтенда.
    """
    return value if value in METHOD_LABELS else ""


def _body(request):
    """
    Безопасно парсит JSON из тела POST-запроса.
    Возвращает кортеж (данные, HttpResponse с ошибкой).
    Если парсинг успешен, ошибка равна None.
    """
    try:
        return json.loads(request.body or "{}"), None
    except ValueError:
        return None, JsonResponse({"error": "неверный формат запроса"}, status=400)


def _pk(value):
    """
    Безопасно преобразует переданное значение (обычно строку) в int для поиска по ID.
    Возвращает None, если преобразование невозможно.
    """
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def login_view(request):
    """
    Отображает форму входа и обрабатывает попытку авторизации.
    Если пользователь уже авторизован, редиректит на главную страницу (method).
    После успешного входа также редиректит на главную.
    """
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
    """
    Разлогинивает пользователя и возвращает его на страницу входа.
    """
    logout(request)
    return redirect("login")


def _filtered(params):
    """
    Строит QuerySet сотрудников на основе параметров фильтрации.
    params — словарь (обычно request.GET), содержащий параметры:
    - q: поиск по табельному номеру или ФИО
    - dep: фильтр по цеху (подразделению)
    - production: фильтр по производству
    - service: фильтр по службе
    - okrug: фильтр по округу (поддерживает 'none' для пустых значений)
    - uik: фильтр по номеру УИК
    - method: фильтр по запланированному способу голосования (поддерживает 'none')
    - where: фильтр по фактическому месту голосования (поддерживает 'none')
    - voted: фильтр по явке ('yes' или 'no')
    """
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

    # Поиск по словам (каждое слово ищется отдельно в tab_number, surname, name, patronymic)
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


def _counts(qs=None):
    """
    Считает агрегированную статистику по базе или переданному QuerySet.
    Возвращает словарь с тремя блоками:
    - method: сколько человек запланировали каждый способ голосования
    - voted: сколько человек фактически проголосовали каждым способом (только те, у кого voted=True)
    - no_uik: количество сотрудников без привязки к УИК
    """
    base = qs if qs is not None else Employee.objects.all()
    agg = base.aggregate(
        deg=Count("id", filter=Q(method=DEG)),
        uik=Count("id", filter=Q(method=UIK)),
        uvz=Count("id", filter=Q(method=UVZ)),
        u19=Count("id", filter=Q(method=UIK19)),
        none=Count("id", filter=Q(method="")),
        total=Count("id"),
    )
    voted = base.filter(voted=True).aggregate(
        deg=Count("id", filter=Q(voted_method=DEG)),
        uik=Count("id", filter=Q(voted_method=UIK)),
        uvz=Count("id", filter=Q(voted_method=UVZ)),
        u19=Count("id", filter=Q(voted_method=UIK19)),
        none=Count("id", filter=Q(voted_method="")),
        total=Count("id"),
    )
    return {
        "method": agg,
        "voted": voted,
        "no_uik": base.filter(uik="").count(),
    }


def _page_window(page, size=5):
    """
    Вычисляет диапазон номеров страниц для пагинатора.
    Возвращает "окно" страниц размером size вокруг текущей страницы.
    Если страниц меньше size, возвращает все страницы.
    """
    total = page.paginator.num_pages
    current = page.number
    if total <= size:
        return range(1, total + 1)
    half = size // 2
    start = current - half
    end = current + half
    # Сдвигаем окно, если выходим за границы
    if start < 1:
        start, end = 1, size
    if end > total:
        start, end = total - size + 1, total
    return range(start, end + 1)


def _context(request):
    """
    Готовит общий контекст для страниц со списками сотрудников (method, elections).
    Включает в себя пагинацию, отфильтрованные данные, счетчики статистики
    и справочники для выпадающих списков фильтров.
    """
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


@login_required
def method_page(request):
    """
    Страница со списком сотрудников для простановки способа голосования.
    """
    return render(request, "method.html", _context(request))


@login_required
def elections_page(request):
    """
    Страница со списком сотрудников для отметки явки (голосования).
    """
    return render(request, "elections.html", _context(request))


@login_required
def upload_page(request):
    """
    Страница загрузки базы сотрудников из Excel-файла.
    Доступна только пользователям с правами оператора (is_operator).
    """
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
    """
    Обработчик POST-запроса с Excel-файлом для импорта/обновления базы сотрудников.
    Вызывает функцию import_base из services.py и сохраняет результат в сессию.
    """
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
    """
    API эндпоинт для обновления способа голосования одного сотрудника.
    Принимает JSON: {"id": ..., "method": ..., "filters": {...}}
    """
    data, bad = _body(request)
    if bad:
        return bad
    changed = Employee.objects.filter(pk=_pk(data.get("id"))).update(
        method=_known_method(data.get("method", ""))
    )
    if not changed:
        return JsonResponse({"error": "работник не найден"}, status=404)
    # Возвращаем обновленную статистику с учетом текущих фильтров
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


MARK_FIELDS = {
    "mark_deg": DEG,
    "mark_uvz": UVZ,
}


@login_required
@require_POST
def api_mark(request):
    data, bad = _body(request)
    if bad:
        return bad
    field = data.get("field")
    if field not in MARK_FIELDS:
        return JsonResponse({"error": "неизвестное поле"}, status=400)
    person = Employee.objects.filter(pk=_pk(data.get("id"))).first()
    if person is None:
        return JsonResponse({"error": "работник не найден"}, status=404)
    if person.method != MARK_FIELDS[field]:
        return JsonResponse(
            {"error": "способ голосования сотрудника не соответсвует этому полю"},
            status=400,
        )
    setattr(person, field, bool(data.get("value")))
    person.save(update_fields=[field])
    return JsonResponse({field: getattr(person, field)})


@login_required
@require_POST
def api_voted(request):
    """
    API эндпоинт для отметки явки одного сотрудника.
    Принимает JSON: {"id": ..., "voted": true/false, "filters": {...}}
    """
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
    """
    API эндпоинт для массовой отметки явки.
    Применяет действие (voted=true/false) ко всем сотрудникам, попадающим под filters.
    Принимает JSON: {"voted": true/false, "filters": {...}}
    """
    data, bad = _body(request)
    if bad:
        return bad
    voted = bool(data.get("voted"))
    filters = data.get("filters") or {}
    target = _filtered(filters)
    skipped = 0
    # При отметке явки пропускаем тех, у кого не выбран способ голосования
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
    """
    API эндпоинт для получения статистики по УИКам (используется в модальном окне).
    Группирует сотрудников по УИК и считает количество людей и явку.
    """
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
    """
    Страница со списком доступных отчётов и формой конструктора сводного отчёта.
    Передает в контекст счетчики цехов/производств и справочники для формы конструктора.
    """
    departments_count = (
        Employee.objects.exclude(department="").values("department").distinct().count()
    )
    productions_count = (
        Employee.objects.exclude(production="").values("service").distinct().count()
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
            # Справочники для формы конструктора сводного отчёта
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


@login_required
def export_summary(request):
    """
    Генерация и отдача Excel-отчета "Сводная таблица по цехам".
    """
    book = summary_table()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    name = f"svodka_po_ceham_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response


@login_required
def export_summary_no_19(request):
    """
    Генерация и отдача Excel-отчета "Сводная таблица по цехам (Без 19 округа)".
    """
    book = summary_table_no_u19()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    name = f"svodka_po_ceham_bez_u19_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response


@login_required
def export_productions(request):
    """
    Генерация и отдача Excel-отчета "Разделение по производствам (Голосование)".
    """
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
    """
    Генерация и отдача полного Excel-отчета со всеми сотрудниками ("Полная таблица").
    """
    book = export_xlsx()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = 'attachment; filename="employees.xlsx"'
    book.save(response)
    return response


def _archive_response(request, mode, prefix):
    """
    Вспомогательная функция для генерации ZIP-архивов с отчетами по каждому цеху.
    mode: 'turnout' (явка) или 'method' (способы голосования)
    prefix: префикс имени файла (yavka или sposob)
    """
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
def export_production_methods(request):
    """
    Генерация и отдача Excel-отчета "Способы голосования по производствам".
    """
    book = production_method_table()
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    name = f"sposoby_po_proizvodstvam_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response


@login_required
def export_production_methods_no_19(request):
    """
    Генерация и отдача Excel-отчета "Способы голосования по производствам (Без 19 округа)".
    """
    book = production_method_table(exclude_u19=True)
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    name = f"sposoby_po_proizvodstvam_bez_u19_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response


@login_required
def export_archive(request):
    """
    Генерация и отдача ZIP-архива с отчетами по явке для каждого цеха.
    """
    return _archive_response(request, "turnout", "yavka")


@login_required
def export_method_archive(request):
    """
    Генерация и отдача ZIP-архива с отчетами по способам голосования для каждого цеха.
    """
    return _archive_response(request, "method", "sposob")


@login_required
def export_custom_report(request):
    """
    Генерация и отдача сводного отчета по сотрудникам на основе фильтров из формы.
    Фильтры передаются через GET-параметры. В зависимости от значения параметра grouping
    формируются разные типы отчёта:
    - "people" (по умолчанию): список по каждому сотруднику (вызывает custom_report)
    - "production_with_depts": агрегация по производствам с разбивкой по цехам
    - "production_without_depts": агрегация по производствам без разбивки по цехам
    """
    grouping = request.GET.get("grouping", "people")

    if grouping == "production_with_depts":
        book = custom_production_summary(request.GET, include_depts=True)
        name = (
            f"svodny_po_proizvodstvam_s_cehami_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
        )
    elif grouping == "production_without_depts":
        book = custom_production_summary(request.GET, include_depts=False)
        name = f"svodny_po_proizvodstvam_{timezone.localtime():%Y%m%d_%H%M}.xlsx"
    else:
        book = custom_report(request.GET)
        name = f"svodny_otchet_{timezone.localtime():%Y%m%d_%H%M}.xlsx"

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = f'attachment; filename="{name}"'
    book.save(response)
    return response
