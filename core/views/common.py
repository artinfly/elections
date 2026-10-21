"""Общее для всех view: роли пользователя, фильтр сотрудников, счётчики, контекст списков."""

from typing import Any, Optional

from django.core.paginator import Paginator
from django.db.models import Count, Q, QuerySet
from django.http import HttpRequest

from ..models import DEG, METHODS, UIK, UIK19, UVZ, Employee

# Количество сотрудников на одной странице таблицы.
PER_PAGE = 100


def is_operator(user: Any) -> bool:
    """
    Проверяет наличие прав оператора или суперпользователя.

    Аргументы:
        user: объект пользователя Django.

    Возвращает:
        bool: True, если пользователь имеет права оператора.
    """
    return user.is_superuser or user.groups.filter(name="operator").exists()


def can_edit(user: Any) -> bool:
    """
    Проверяет право пользователя на редактирование данных.

    Аргументы:
        user: объект пользователя Django.

    Возвращает:
        bool: True, если пользователю разрешено вносить изменения.
    """
    return not user.groups.filter(name="viewer").exists()


def is_uik_uvz_viewer(user: Any) -> bool:
    """
    Проверяет роль «наблюдатель УИК-УВЗ» (группа uik_uvz_viewer).

    Такому пользователю доступна только страница явки, и в ней видны
    лишь сотрудники, выбравшие или проголосовавшие на участке УИК-УВЗ.
    """
    return user.groups.filter(name="uik_uvz_viewer").exists()


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


def filter_options() -> dict[str, list[str]]:
    """
    Получает уникальные значения для всех выпадающих списков фильтров.

    Описание:
        Выполняет оптимизированные SQL-запросы с DISTINCT для получения
        уникальных значений. Это предотвращает выгрузку всей таблицы в память Python,
        что было бы критично для больших баз данных.

    Возвращает:
        dict: словари с отсортированными списками уникальных значений.
    """
    employees = Employee.objects.values(
        "department", "production", "service", "okrug", "uik"
    )

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


def filter_employees(params: dict) -> QuerySet:
    """
    Строит выборку (QuerySet) сотрудников на основе переданных фильтров.

    Аргументы:
        params: словарь с параметрами фильтрации.

    Возвращает:
        QuerySet: отфильтрованный список сотрудников.
    """
    qs = Employee.objects.all()

    # Гибкий поиск по ФИО и табельному номеру.
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


def vote_counts(qs: Optional[QuerySet] = None) -> dict[str, Any]:
    """
    Считает агрегированную статистику одним SQL-запросом.

    Описание:
        Возвращает количество людей по каждому способу голосования (план)
        и по фактической явке. Используется для обновления счетчиков на фронтенде.

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
        plan_none=Count("id", filter=Q(method="", absence=False)),
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
        все страницы в подвале таблицы, а только ближайшие.

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


def list_context(request: HttpRequest) -> dict[str, Any]:
    """
    Готовит общий контекст для рендеринга HTML-страниц со списками.

    Аргументы:
        request: HTTP-запрос.

    Возвращает:
        dict: контекст для передачи в функцию render().
    """
    qs = filter_employees(request.GET)
    page = Paginator(qs, PER_PAGE).get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)

    options = filter_options()

    return {
        "page": page,
        "rows": page.object_list,
        "found": page.paginator.count,
        "counts": vote_counts(qs),
        "is_operator": is_operator(request.user),
        "methods": METHODS,
        "departments": options["departments"],
        "productions": options["productions"],
        "services": options["services"],
        "okrugs": options["okrugs"],
        "uiks": options["uiks"],
        "f": request.GET,
        "query": params.urlencode(),
        "page_range": _page_window(page),
        "can_edit": can_edit(request.user),
        "is_uik_uvz_viewer": is_uik_uvz_viewer(request.user),
    }
