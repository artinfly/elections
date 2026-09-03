"""
Модуль контроллеров (Views).

Описание:
    Принимает HTTP-запросы от браузера, проверяет права доступа пользователя,
    обращается к базе данных через модели и возвращает либо HTML-страницу,
    либо скачиваемый файл (Excel/ZIP), либо JSON-ответ для асинхронных кнопок.
"""

import json
from typing import Any, Optional

from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q, QuerySet
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
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

# Количество сотрудников на одной странице таблицы.
PER_PAGE = 100

# Варианты фильтрации по избирательным округам для конструктора отчетов.
CUSTOM_OKRUG_OPTIONS = [
    ("", "Все"),
    ("none", "Пусто"),
    ("19", "19"),
    ("20", "20"),
    ("21", "21"),
    ("20+21", "20+21"),
]

# Связь полей отметок в базе с кодами способов голосования.
# Используется для проверки: нельзя поставить отметку о регистрации на ДЭГ,
# если сотрудник выбрал голосование на обычном участке.
MARK_FIELDS = {"mark_deg": DEG, "mark_uvz": UVZ}


# ==============================================================================
# Вспомогательные функции (Helpers)
# ==============================================================================


def is_operator(user: Any) -> bool:
    """
    Проверяет наличие прав оператора или суперпользователя.

    Описание:
        Операторы имеют доступ к загрузке файлов и административным функциям.
        Суперпользователь имеет все права по умолчанию.

    Аргументы:
        user: объект пользователя Django.

    Возвращает:
        bool: True, если пользователь имеет права оператора.
    """
    return user.is_superuser or user.groups.filter(name="operator").exists()


def can_edit(user: Any) -> bool:
    """
    Проверяет право пользователя на редактирование данных.

    Описание:
        Все авторизованные пользователи могут редактировать данные,
        кроме тех, кто явно добавлен в группу 'viewer' (только просмотр).

    Аргументы:
        user: объект пользователя Django.

    Возвращает:
        bool: True, если пользователю разрешено вносить изменения.
    """
    return not user.groups.filter(name="viewer").exists()


def _known_method(value: Any) -> str:
    """
    Валидирует код способа голосования.

    Описание:
        Проверяет, что переданное значение является одним из зарегистрированных
        кодов (deg, uik, uvz, u19). Если передан мусор — возвращает пустую строку,
        чтобы избежать записи невалидных данных в базу.

    Аргументы:
        value: код способа голосования.

    Возвращает:
        str: Валидный код способа или пустая строка.
    """
    return str(value) if value in METHOD_LABELS else ""


def _parse_json_body(
    request: HttpRequest,
) -> tuple[Optional[dict], Optional[JsonResponse]]:
    """
    Безопасно извлекает JSON из тела POST-запроса.

    Описание:
        Защищает сервер от падения с ошибкой 500, если фронтенд прислал
        битый или пустой JSON. В случае ошибки парсинга сразу формирует ответ.

    Аргументы:
        request: HTTP-запрос.

    Возвращает:
        tuple: (словарь с данными, объект ошибки). Если ошибка есть, её нужно вернуть клиенту.
    """
    try:
        return json.loads(request.body or "{}"), None
    except json.JSONDecodeError:
        return None, JsonResponse({"error": "неверный формат запроса"}, status=400)


def _safe_int(value: Any) -> Optional[int]:
    """
    Безопасно преобразует значение в целое число.

    Описание:
        Используется для чтения ID из JSON. Если фронтенд передал строку или null,
        функция вернет None, что позволит корректно обработать ошибку 400.

    Аргументы:
        value: значение для преобразования.

    Возвращает:
        int или None: число при успехе, None при ошибке.
    """
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _clean(params: dict, key: str) -> str:
    """
    Безопасно извлекает строковый параметр из словаря.

    Описание:
        URL-параметры всегда приходят как строки, но JSON-фильтры из API могут
        содержать числа или списки. Функция приводит всё к строке и удаляет
        пробелы по краям. Нестроковые значения считаются пустыми.

    Аргументы:
        params: словарь параметров (request.GET или JSON).
        key: имя параметра.

    Возвращает:
        str: очищенная строка или пустая строка.
    """
    value = params.get(key)
    return value.strip() if isinstance(value, str) else ""


def _make_excel_response(workbook: Any, filename_prefix: str) -> HttpResponse:
    """
    Формирует HTTP-ответ для скачивания Excel-файла браузером.

    Описание:
        Устанавливает правильные заголовки, чтобы браузер понял, что нужно
        скачать файл, а не пытаться отобразить его на странице.
        В имя файла добавляется текущая дата и время.

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


def _validate_excel_file(file: Any) -> None:
    """
    Проверяет расширение загружаемого файла.

    Описание:
        Система принимает только файлы формата Excel (.xlsx).
        Попытка загрузить CSV или картинку будет пресечена.

    Аргументы:
        file: объект загруженного файла из request.FILES.

    Исключения:
        ValueError: если файл отсутствует или имеет неверное расширение.
    """
    if not file or not file.name.lower().endswith(".xlsx"):
        raise ValueError("Ошибка: принимается только формат .xlsx")


def _get_filter_options() -> dict[str, list[str]]:
    """
    Получает уникальные значения для всех выпадающих списков фильтров.

    Описание:
        Делает один широкий запрос к базе, чтобы собрать все уникальные цеха,
        производства и УИКи. Это работает быстрее, чем делать 5 отдельных
        SQL-запросов (оптимизация N+1).

    Возвращает:
        dict: словари с отсортированными списками уникальных значений.
    """
    # Один запрос получает все нужные поля сразу.
    employees = Employee.objects.values(
        "department", "production", "service", "okrug", "uik"
    )

    # Извлекаем уникальные значения средствами Python (быстрее SQL DISTINCT для малых объемов).
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
    Строит выборку (QuerySet) сотрудников на основе переданных фильтров.

    Описание:
        Применяет фильтры из URL-параметров или JSON. Поддерживает поиск
        по частям ФИО и специфичные фильтры (например, "Пусто" для округа).

    Аргументы:
        params: словарь с параметрами фильтрации.

    Возвращает:
        QuerySet: отфильтрованный список сотрудников.
    """
    qs = Employee.objects.all()

    # Гибкий поиск по ФИО и табельному номеру.
    # Разбиваем фразу на слова: поиск "иван иванов" найдет "Иванов Иван Иванович".
    search = _clean(params, "q")
    for part in search.split():
        qs = qs.filter(
            Q(tab_number__icontains=part)
            | Q(surname__icontains=part)
            | Q(name__icontains=part)
            | Q(patronymic__icontains=part)
        )

    # Простые фильтры (точное совпадение).
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

    # Специфичные фильтры с поддержкой значения "none" (Пусто).
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
    Считает агрегированную статистику одним SQL-запросом.

    Описание:
        Возвращает количество людей по каждому способу голосования (план)
        и по фактической явке. Используется для обновления счетчиков (плашек)
        в верхней части страницы без перезагрузки.

    Аргументы:
        qs: QuerySet для подсчета (по умолчанию все сотрудники).

    Возвращает:
        dict: структура данных со статистикой.
    """
    base = qs if qs is not None else Employee.objects.all()
    agg = base.aggregate(
        # План по способам голосования.
        plan_deg=Count("id", filter=Q(method=DEG)),
        plan_uik=Count("id", filter=Q(method=UIK)),
        plan_uvz=Count("id", filter=Q(method=UVZ)),
        plan_u19=Count("id", filter=Q(method=UIK19)),
        plan_none=Count("id", filter=Q(method="")),
        plan_total=Count("id"),
        # Фактическая явка (только те, у кого voted=True).
        voted_deg=Count("id", filter=Q(voted=True, voted_method=DEG)),
        voted_uik=Count("id", filter=Q(voted=True, voted_method=UIK)),
        voted_uvz=Count("id", filter=Q(voted=True, voted_method=UVZ)),
        voted_u19=Count("id", filter=Q(voted=True, voted_method=UIK19)),
        voted_none=Count("id", filter=Q(voted=True, voted_method="")),
        voted_total=Count("id", filter=Q(voted=True)),
        # Служебная статистика.
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

    Описание:
        Формирует "окно" из нескольких страниц вокруг текущей, чтобы не показывать
        все 100 страниц в подвале таблицы, а только ближайшие (например, 3, 4, [5], 6, 7).

    Аргументы:
        page: объект Page из Django Paginator.
        size: количество страниц в окне.

    Возвращает:
        range: диапазон номеров страниц.
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
    Готовит общий контекст для рендеринга HTML-страниц со списками.

    Описание:
        Собирает все данные, нужные шаблонам: текущую страницу, фильтры,
        статистику и списки для выпадающих меню.

    Аргументы:
        request: HTTP-запрос.

    Возвращает:
        dict: контекст для передачи в функцию render().
    """
    qs = _filtered(request.GET)
    page = Paginator(qs, PER_PAGE).get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)

    # Получаем опции фильтров одним запросом (оптимизация N+1).
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
    Главная страница: список сотрудников для простановки способа голосования.
    """
    return render(request, "method.html", _context(request))


@login_required
def elections_page(request: HttpRequest) -> HttpResponse:
    """
    Страница выборов: список сотрудников для отметки фактической явки.
    """
    return render(request, "elections.html", _context(request))


@login_required
def upload_page(request: HttpRequest) -> HttpResponse:
    """
    Страница загрузки файлов из Excel (доступна только операторам).
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
    Страница экспорта: список готовых отчетов и конструктор кастомных сводок.
    """
    filter_options = _get_filter_options()

    return render(
        request,
        "export.html",
        {
            "counts": _counts(),
            "is_operator": is_operator(request.user),
            "departments_count": len(filter_options["departments"]),
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
    Обработчик формы загрузки основного списка сотрудников.
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
    Обработчик формы загрузки отчета штаба (обновление способов голосования).
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
# API endpoints (Асинхронные запросы без перезагрузки страницы)
# ==============================================================================


@login_required
@require_POST
def api_method(request: HttpRequest) -> JsonResponse:
    """
    API: обновление запланированного способа голосования для одного сотрудника.
    """
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

    # Возвращаем обновленную статистику для перерисовки плашек на фронте.
    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_mark(request: HttpRequest) -> JsonResponse:
    """
    API: простановка отметок регистрации (ДЭГ или УВЗ).
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

    employee = get_object_or_404(Employee.objects, pk=employee_id)

    # Проверка бизнес-логики: отметка должна соответствовать выбранному способу.
    if employee.method != MARK_FIELDS[field]:
        return JsonResponse({"error": "способ не соответствует полю"}, status=400)

    setattr(employee, field, bool(data.get("value")))
    employee.save(update_fields=[field])

    return JsonResponse({field: getattr(employee, field)})


@login_required
@require_POST
def api_voted(request: HttpRequest) -> JsonResponse:
    """
    API: отметка фактической явки для одного сотрудника.
    """
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    data, bad = _parse_json_body(request)
    if bad:
        return bad

    employee_id = _safe_int(data.get("id"))
    if employee_id is None:
        return JsonResponse({"error": "некорректный ID"}, status=400)

    employee = get_object_or_404(Employee.objects, pk=employee_id)

    voted = bool(data.get("voted"))
    # Нельзя отметить явку, если сотрудник еще не выбрал способ голосования.
    if voted and not employee.method:
        return JsonResponse({"error": "Не выбран способ голосования"}, status=400)

    mark_voted([employee.tab_number], voted=voted)

    return JsonResponse(_counts(_filtered(data.get("filters") or {})))


@login_required
@require_POST
def api_bulk_voted(request: HttpRequest) -> JsonResponse:
    """
    API: массовая отметка явки по текущим фильтрам.
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
        # Пропускаем сотрудников без выбранного способа, чтобы не сломать логику.
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
@require_POST
def api_toggle_absence(request: HttpRequest, employee_id: int) -> JsonResponse:
    """
    API: переключение отметки «Отсутствие по УП» (уважительная причина).

    Описание:
        Инвертирует текущее булево значение поля absence у конкретного сотрудника.
        Используется для кнопки на странице «Способ голосования», которая должна
        срабатывать без перезагрузки страницы.

    Аргументы:
        request: HTTP-запрос.
        employee_id: уникальный идентификатор сотрудника из URL.

    Возвращает:
        JsonResponse с ID сотрудника и новым статусом отсутствия.
    """
    # Проверка прав: только пользователи с правами редактирования могут менять отметки.
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    # Ищем сотрудника. Если ID не существует, Django автоматически вернет 404.
    employee = get_object_or_404(Employee.objects, id=employee_id)

    # Инвертируем флаг (True -> False, False -> True).
    employee.absence = not employee.absence

    # Сохраняем только измененное поле, чтобы не делать лишних записей в БД.
    employee.save(update_fields=["absence"])

    return JsonResponse({"id": employee.id, "absence": employee.absence})


# ==============================================================================
# Экспорт отчетов (Standard)
# ==============================================================================


@login_required
def export_summary(request: HttpRequest) -> HttpResponse:
    """Экспорт сводной таблицы по цехам."""
    return _make_excel_response(summary_table(), "svodka_po_ceham")


@login_required
def export_summary_no_19(request: HttpRequest) -> HttpResponse:
    """Экспорт сводной таблицы по цехам без учета 19 округа."""
    return _make_excel_response(summary_table_no_u19(), "svodka_po_ceham_bez_u19")


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


def _archive_response(request: HttpRequest, mode: str, prefix: str) -> HttpResponse:
    """
    Внутренний хелпер для генерации ZIP-архивов с отчетами по цехам.
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
    """Экспорт архива отчетов по явке."""
    return _archive_response(request, "turnout", "yavka")


@login_required
def export_custom_archive(request: HttpRequest) -> HttpResponse:
    """Экспорт кастомного архива отчетов по цехам."""
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
    """Экспорт архива отчетов по способам голосования."""
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
    """
    Страница входа в систему (форма логина и пароля).
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
    Завершение сеанса пользователя и выход из системы.
    """
    logout(request)
    return redirect("login")
