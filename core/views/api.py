"""JSON API для кнопок на страницах: отметки способа, регистрации и явки без перезагрузки."""

import json
from typing import Any

from django.contrib.auth.decorators import login_required
from django.db.models import Count, Q
from django.http import HttpRequest, JsonResponse
from django.shortcuts import get_object_or_404
from django.views.decorators.http import require_POST

from ..models import DEG, METHOD_LABELS, UIK, UVZ, Employee
from ..services.importing import mark_voted, set_turnout
from .common import can_edit, filter_employees, vote_counts

# Поле отметки регистрации -> способ, при котором её можно поставить.
MARK_FIELDS = {"mark_deg": DEG, "mark_uvz": UVZ}


def _known_method(value: Any) -> str:
    """
    Валидирует код способа голосования.

    Описание:
        Проверяет, что переданное значение является одним из зарегистрированных
        кодов. Если передан мусор — возвращает пустую строку, чтобы избежать
        записи невалидных данных в базу.

    Аргументы:
        value: код способа голосования.

    Возвращает:
        str: Валидный код способа или пустая строка.
    """
    return str(value) if value in METHOD_LABELS else ""


def _parse_json_body(
    request: HttpRequest,
) -> tuple[dict | None, JsonResponse | None]:
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


def _safe_int(value: Any) -> int | None:
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


@login_required
@require_POST
def api_method(request: HttpRequest) -> JsonResponse:
    """API: обновление запланированного способа голосования для одного сотрудника."""
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

    return JsonResponse(vote_counts(filter_employees(data.get("filters") or {})))


@login_required
@require_POST
def api_mark(request: HttpRequest) -> JsonResponse:
    """API: простановка отметок регистрации (ДЭГ или УВЗ)."""
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

    if employee.method != MARK_FIELDS[field]:
        return JsonResponse({"error": "способ не соответствует полю"}, status=400)

    setattr(employee, field, bool(data.get("value")))
    employee.save(update_fields=[field])

    return JsonResponse({field: getattr(employee, field)})


@login_required
@require_POST
def api_voted(request: HttpRequest) -> JsonResponse:
    """API: отметка фактической явки для одного сотрудника."""
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
    if voted and not employee.method:
        # Способ не выбран — при отметке явки считаем, что голосовал на своём участке
        Employee.objects.filter(tab_number=employee.tab_number).update(method=UIK)

    mark_voted([employee.tab_number], voted=voted)

    return JsonResponse(vote_counts(filter_employees(data.get("filters") or {})))


@login_required
@require_POST
def api_bulk_voted(request: HttpRequest) -> JsonResponse:
    """API: массовая отметка явки по текущим фильтрам."""
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    data, bad = _parse_json_body(request)
    if bad:
        return bad

    voted = bool(data.get("voted"))
    filters = data.get("filters") or {}
    target = filter_employees(filters)

    skipped = 0
    if voted:
        skipped = target.filter(method="").count()
        target = target.exclude(method="")

    changed = set_turnout(target, voted)
    result = vote_counts(filter_employees(filters))
    result["changed"] = changed
    result["skipped"] = skipped

    return JsonResponse(result)


@login_required
def api_uik_stats(request: HttpRequest) -> JsonResponse:
    """API: статистика по УИКам для модального окна."""
    rows = (
        filter_employees(request.GET)
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

    Аргументы:
        request: HTTP-запрос.
        employee_id: уникальный идентификатор сотрудника из URL.

    Возвращает:
        JsonResponse с ID сотрудника и новым статусом отсутствия.
    """
    if not can_edit(request.user):
        return JsonResponse({"error": "нет прав"}, status=403)

    employee = get_object_or_404(Employee.objects, id=employee_id)
    employee.absence = not employee.absence
    employee.save(update_fields=["absence"])

    return JsonResponse({"id": employee.id, "absence": employee.absence})
