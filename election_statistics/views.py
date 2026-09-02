"""
Модуль представлений (Views).

Обрабатывает HTTP-запросы, управляет доступом, рендерит страницы
и отдает ответы с данными или файлами.
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
    custom_reports_archive,
    export_xlsx,
    production_method_table,
    production_table,
    reports_archive,
    summary_table,
    summary_table_no_u19,
)

# ==============================================================================
# Константы
# ==============================================================================

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
    """
    Проверяет, является ли пользователь оператором или суперюзером.

    Аргументы:
        user: объект пользователя Django

    Возвращает:
        True если пользователь имеет права оператора
    """
    return user.is_superuser or user.groups.filter(name="operator").exists()


def can_edit(user: Any) -> bool:
    """
    Проверяет, может ли пользователь редактировать данные.

    Аргументы:
        user: объект пользователя Django

    Возвращает:
        True если пользователь не в группе viewer
    """
    return not user.groups.filter(name="viewer").exists()


def _known_method(value: Any) -> str:
    """
    Валидирует значение способа голосования.

    Аргументы:
        value: код способа голосования

    Возвращает:
        Код способа если валиден, иначе пустая строка
    """
    return str(value) if value in METHOD_LABELS else ""


def _parse_json_body(
    request: HttpRequest,
) -> tuple[Optional[dict], Optional[JsonResponse]]:
    """
    Безопасно парсит JSON из тела POST-запроса.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        Кортеж (данные, ошибка). Если ошибка не None — вернуть её как ответ.
    """
    try:
        return json.loads(request.body or "{}"), None
    except json.JSONDecodeError:
        return None, JsonResponse({"error": "неверный формат запроса"}, status=400)


def _safe_int(value: Any) -> Optional[int]:
    """
    Безопасно преобразует значение в int.

    Аргументы:
        value: значение для преобразования

    Возвращает:
        Целое число или None при ошибке
    """
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _clean(params: dict, key: str) -> str:
    """
    Безопасно достаёт строковый параметр из словаря.

    request.GET всегда отдаёт строки, но JSON-фильтры из API могут
    содержать что угодно (числа, списки) — всё, кроме str, считаем
    пустым значением, чтобы не падать с 500.

    Аргументы:
        params: словарь параметров (request.GET или JSON-фильтры)
        key: имя параметра

    Возвращает:
        Строку без пробелов по краям или ""
    """
    value = params.get(key)
    return value.strip() if isinstance(value, str) else ""


def _make_excel_response(workbook: Any, filename_prefix: str) -> HttpResponse:
    """
    Создает HTTP-ответ с Excel-файлом для скачивания.

    Аргументы:
        workbook: объект openpyxl Workbook
        filename_prefix: префикс имени файла

    Возвращает:
        HttpResponse с Excel-файлом
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


def _validate_excel_file(file: Any) -> None:
    """
    Проверяет, что загруженный файл имеет расширение .xlsx.

    Аргументы:
        file: загруженный файл из request.FILES

    Исключения:
        ValueError: если файл не Excel или отсутствует
    """
    if not file or not file.name.lower().endswith(".xlsx"):
        raise ValueError("Ошибка: принимается только формат .xlsx")


def _get_filter_options() -> dict[str, list[str]]:
    """
    Получает уникальные значения для всех фильтров одним запросом к БД.

    Возвращает:
        Словарь с отсортированными списками уникальных значений
    """
    # Один запрос получает все данные вместо 5 отдельных
    employees = Employee.objects.values(
        "department", "production", "service", "okrug", "uik"
    )

    # Извлекаем уникальные значения в Python (быстрее чем 5 SQL запросов)
    departments = sorted({e["department"] for e in employees if e["department"]})
    productions = sorted({e["production"] for e in employees if e["production"]})
    services = sorted({e["service"] for e in employees if e["service"]})
    okrugs = sorted({e["okrug"] for e in employees if e["okrug"]})
    uiks = sorted({e["uik"] for e in employees if e["uik"]})

    return {
        "departments": departments,
        "productions": productions,
        "services": services,
        "okrugs": okrugs,
        "uiks": uiks,
    }


# ==============================================================================
# Фильтрация и статистика
# ==============================================================================


def _filtered(params: dict) -> QuerySet:
    """
    Строит QuerySet сотрудников на основе параметров фильтрации.

    Аргументы:
        params: словарь с параметрами из request.GET или JSON-фильтров

    Возвращает:
        Отфильтрованный QuerySet сотрудников
    """
    qs = Employee.objects.all()

    # Поиск по ФИО и табельному номеру (разбиваем на слова для гибкого поиска)
    search = _clean(params, "q")
    for part in search.split():
        qs = qs.filter(
            Q(tab_number__icontains=part)
            | Q(surname__icontains=part)
            | Q(name__icontains=part)
            | Q(patronymic__icontains=part)
        )

    # Простые фильтры: маппинг параметр -> поле модели
    filter_mapping = {
        "dep": "department",
        "production": "production",
        "service": "service",
        "uik": "uik",
    }

    for param, field in filter_mapping.items():
        value = _clean(params, param)
        if value:
            qs = qs.filter(**{field: value})

    # Фильтры со специальной логикой (пустое значение как "none")
    okrug = _clean(params, "okrug")
    if okrug == "none":
        qs = qs.filter(okrug="")
    elif okrug:
        qs = qs.filter(okrug=okrug)

    method = _clean(params, "method")
    if method == "none":
        qs = qs.filter(method="")
    elif method:
        qs = qs.filter(method=method)

    where = _clean(params, "where")
    if where == "none":
        qs = qs.filter(voted_method="")
    elif where:
        qs = qs.filter(voted_method=where)

    voted = _clean(params, "voted")
    if voted == "yes":
        qs = qs.filter(voted=True)
    elif voted == "no":
        qs = qs.filter(voted=False)

    return qs


def _counts(qs: Optional[QuerySet] = None) -> dict[str, Any]:
    """
    Считает агрегированную статистику одним запросом к БД.

    Аргументы:
        qs: QuerySet для подсчета (по умолчанию все сотрудники)

    Возвращает:
        Словарь со статистикой по способам голосования и явке
    """
    base = qs if qs is not None else Employee.objects.all()
    agg = base.aggregate(
        # План по способам голосования
        plan_deg=Count("id", filter=Q(method=DEG)),
        plan_uik=Count("id", filter=Q(method=UIK)),
        plan_uvz=Count("id", filter=Q(method=UVZ)),
        plan_u19=Count("id", filter=Q(method=UIK19)),
        plan_none=Count("id", filter=Q(method="")),
        plan_total=Count("id"),
        # Фактическая явка
        voted_deg=Count("id", filter=Q(voted=True, voted_method=DEG)),
        voted_uik=Count("id", filter=Q(voted=True, voted_method=UIK)),
        voted_uvz=Count("id", filter=Q(voted=True, voted_method=UVZ)),
        voted_u19=Count("id", filter=Q(voted=True, voted_method=UIK19)),
        voted_none=Count("id", filter=Q(voted=True, voted_method="")),
        voted_total=Count("id", filter=Q(voted=True)),
        # Сотрудники без УИК
        no_uik=Count("id", filter=Q(uik="")),
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
        "no_uik": agg["no_uik"],
    }


def _page_window(page: Any, size: int = 5) -> range:
    """
    Вычисляет диапазон номеров страниц для пагинатора.

    Аргументы:
        page: объект Page из Django Paginator
        size: количество страниц в окне (по умолчанию 5)

    Возвращает:
        range с номерами страниц для отображения
    """
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
    """
    Готовит общий контекст для страниц со списками сотрудников.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        Словарь контекста для шаблона
    """
    qs = _filtered(request.GET)
    page = Paginator(qs, PER_PAGE).get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)

    # Получаем все опции фильтров одним запросом (оптимизация N+1)
    filter_options = _get_filter_options()

    return {
        "page": page,
        "rows": page.object_list,
        "found": page.paginator.count,
        "counts": _counts(qs),
        "is_operator": is_operator(request.user),
        "methods": METHODS,
        "departments": filter_options["departments"],
        "productions": filter_options["productions"],
        "services": filter_options["services"],
        "okrugs": filter_options["okrugs"],
        "uiks": filter_options["uiks"],
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
    """
    Страница со списком сотрудников для простановки способа голосования.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        HttpResponse с отрендеренным шаблоном
    """
    return render(request, "method.html", _context(request))


@login_required
def elections_page(request: HttpRequest) -> HttpResponse:
    """
    Страница со списком сотрудников для отметки явки.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        HttpResponse с отрендеренным шаблоном
    """
    return render(request, "elections.html", _context(request))


@login_required
def upload_page(request: HttpRequest) -> HttpResponse:
    """
    Страница загрузки файлов (доступна только операторам).

    Аргументы:
        request: HTTP запрос

    Возвращает:
        HttpResponse с отрендеренным шаблоном или страницей отказа в доступе
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
def export_page(request: HttpRequest) -> HttpResponse:
    """
    Страница со списком доступных отчетов и формой конструктора.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        HttpResponse с отрендеренным шаблоном
    """
    filter_options = _get_filter_options()

    return render(
        request,
        "export.html",
        {
            "counts": _counts(),
            "is_operator": is_operator(request.user),
            # Считаем количеством в уже полученных списках — без лишних запросов
            "departments_count": len(filter_options["departments"]),
            # «Производства» отчётов живут в поле service (см. комментарий
            # в reports.py) — считаем уникальные непустые service
            "productions_count": len(filter_options["services"]),
            "msg": request.session.pop("msg", ""),
            "productions": filter_options["productions"],
            "services": filter_options["services"],
            "departments": filter_options["departments"],
            "methods": METHODS,
            "uiks": filter_options["uiks"],
            "custom_okrugs": CUSTOM_OKRUG_OPTIONS,
        },
    )


# ==============================================================================
# Обработчики загрузки (Upload handlers)
# ==============================================================================


@login_required
@require_POST
def upload_base(request: HttpRequest) -> HttpResponse:
    """
    Обработчик загрузки основной базы сотрудников.

    Аргументы:
        request: HTTP запрос с файлом

    Возвращает:
        Редирект на страницу загрузки с сообщением в сессии
    """
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

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
def upload_voting_choices(request: HttpRequest) -> HttpResponse:
    """
    Обработчик загрузки отчета штаба для обновления способов голосования.

    Аргументы:
        request: HTTP запрос с файлом

    Возвращает:
        Редирект на страницу загрузки с сообщением в сессии
    """
    if not is_operator(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    upload = request.FILES.get("file")
    try:
        _validate_excel_file(upload)
    except ValueError as exc:
        request.session["msg"] = str(exc)
        return redirect("upload")

    try:
        updated, total, errors = import_voting_choices(upload)
        request.session["msg"] = (
            f"Обработано строк: {total}. Обновлено способов: {updated}, "
            f"ошибок/пропусков: {errors}"
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
    """
    API: обновление способа голосования для одного сотрудника.

    Аргументы:
        request: POST запрос с JSON {id, method, filters}

    Возвращает:
        JSON с обновленной статистикой или ошибкой (400/403/404)
    """
    # Viewer может только смотреть: проверяем право на сервере, а не в шаблоне
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    data, bad = _parse_json_body(request)
    if bad:
        return bad

    employee_id = _safe_int(data.get("id"))
    if employee_id is None:
        return JsonResponse({"error": "некорректный ID"}, status=400)

    changed = Employee.objects.filter(pk=employee_id).update(
        method=_known_method(data.get("method", ""))
    )
    if not changed:
        return JsonResponse({"error": "работник не найден"}, status=404)

    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_mark(request: HttpRequest) -> JsonResponse:
    """
    API: простановка отметок регистрации (mark_deg, mark_uvz).

    Аргументы:
        request: POST запрос с JSON {id, field, value}

    Возвращает:
        JSON с новым значением поля или ошибкой (400/403/404)
    """
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    data, bad = _parse_json_body(request)
    if bad:
        return bad

    employee_id = _safe_int(data.get("id"))
    if employee_id is None:
        return JsonResponse({"error": "некорректный ID"}, status=400)

    field = data.get("field")
    if field not in MARK_FIELDS:
        return JsonResponse({"error": "неизвестное поле"}, status=400)

    person = Employee.objects.filter(pk=employee_id).first()
    if person is None:
        return JsonResponse({"error": "работник не найден"}, status=404)

    # Проверка бизнес-логики: отметка должна соответствовать способу
    if person.method != MARK_FIELDS[field]:
        return JsonResponse({"error": "способ не соответствует полю"}, status=400)

    setattr(person, field, bool(data.get("value")))
    person.save(update_fields=[field])

    return JsonResponse({field: getattr(person, field)})


@login_required
@require_POST
def api_voted(request: HttpRequest) -> JsonResponse:
    """
    API: отметка явки для одного сотрудника.

    Аргументы:
        request: POST запрос с JSON {id, voted, filters}

    Возвращает:
        JSON с обновленной статистикой или ошибкой (400/403/404)
    """
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    data, bad = _parse_json_body(request)
    if bad:
        return bad

    employee_id = _safe_int(data.get("id"))
    if employee_id is None:
        return JsonResponse({"error": "некорректный ID"}, status=400)

    person = Employee.objects.filter(pk=employee_id).first()
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
    """
    API: массовая отметка явки по текущим фильтрам.

    Аргументы:
        request: POST запрос с JSON {voted, filters}

    Возвращает:
        JSON со статистикой и количеством измененных/пропущенных или ошибкой
    """
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    data, bad = _parse_json_body(request)
    if bad:
        return bad

    voted = bool(data.get("voted"))
    filters = data.get("filters") or {}
    target = _filtered(filters)

    skipped = 0
    if voted:
        # Пропускаем сотрудников без выбранного способа
        skipped = target.filter(method="").count()
        target = target.exclude(method="")

    changed = set_turnout(target, voted)
    result = _counts(_filtered(filters))
    result["changed"] = changed
    result["skipped"] = skipped

    return JsonResponse(result)


@login_required
def api_uik_stats(request: HttpRequest) -> JsonResponse:
    """
    API: статистика по УИКам для модального окна.

    Аргументы:
        request: GET запрос с параметрами фильтрации

    Возвращает:
        JSON с массивом статистики по каждому УИК
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


# ==============================================================================
# Экспорт отчетов (Standard)
# ==============================================================================


@login_required
def export_summary(request: HttpRequest) -> HttpResponse:
    """
    Экспорт сводной таблицы по цехам.

    Возвращает:
        HttpResponse с Excel-файлом
    """
    return _make_excel_response(summary_table(), "svodka_po_ceham")


@login_required
def export_summary_no_19(request: HttpRequest) -> HttpResponse:
    """
    Экспорт сводной таблицы по цехам без учета 19 округа.

    Возвращает:
        HttpResponse с Excel-файлом
    """
    return _make_excel_response(summary_table_no_u19(), "svodka_po_ceham_bez_u19")


@login_required
def export_productions(request: HttpRequest) -> HttpResponse:
    """
    Экспорт таблицы по производствам.

    Возвращает:
        HttpResponse с Excel-файлом
    """
    return _make_excel_response(production_table(), "po_proizvodstvam")


@login_required
def export_production_methods(request: HttpRequest) -> HttpResponse:
    """
    Экспорт таблицы способов голосования по производствам.

    Возвращает:
        HttpResponse с Excel-файлом
    """
    return _make_excel_response(production_method_table(), "sposoby_po_proizvodstvam")


@login_required
def export_production_methods_no_19(request: HttpRequest) -> HttpResponse:
    """
    Экспорт таблицы способов голосования по производствам без 19 округа.

    Возвращает:
        HttpResponse с Excel-файлом
    """
    return _make_excel_response(
        production_method_table(exclude_u19=True), "sposoby_po_proizvodstvam_bez_u19"
    )


@login_required
def export_employees(request: HttpRequest) -> HttpResponse:
    """
    Экспорт полной таблицы сотрудников.

    Возвращает:
        HttpResponse с Excel-файлом
    """
    return _make_excel_response(export_xlsx(), "employees")


def _archive_response(request: HttpRequest, mode: str, prefix: str) -> HttpResponse:
    """
    Хелпер для генерации ZIP-архивов с отчетами по цехам.

    Аргументы:
        request: HTTP запрос
        mode: режим отчета ("turnout" или "method")
        prefix: префикс имени файла

    Возвращает:
        HttpResponse с ZIP-файлом или редирект с ошибкой
    """
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
    """
    Экспорт архива отчетов по явке.

    Возвращает:
        HttpResponse с ZIP-файлом
    """
    return _archive_response(request, "turnout", "yavka")


@login_required
def export_custom_archive(request: HttpRequest) -> HttpResponse:
    """
    Экспорт кастомного архива отчетов по цехам.

    Возвращает:
        HttpResponse с ZIP-файлом или редирект с ошибкой
    """
    moment = timezone.localtime()
    archiver = custom_reports_archive(request.GET, moment)

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
    """
    Экспорт архива отчетов по способам голосования.

    Возвращает:
        HttpResponse с ZIP-файлом
    """
    return _archive_response(request, "method", "sposob")


@login_required
def export_custom_report(request: HttpRequest) -> HttpResponse:
    """
    Экспорт кастомного отчета (по людям или по производствам).

    Аргументы:
        request: GET запрос с параметром grouping

    Возвращает:
        HttpResponse с Excel-файлом
    """
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
    """
    Страница входа в систему.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        HttpResponse с формой входа или редирект для авторизованных
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


@login_required
def logout_view(request: HttpRequest) -> HttpResponse:
    """
    Выход из системы.

    Аргументы:
        request: HTTP запрос

    Возвращает:
        Редирект на страницу входа
    """
    logout(request)
    return redirect("login")
