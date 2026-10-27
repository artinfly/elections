"""JSON API для отдельного фронтенда (адреса /api/…, описание — docs/api.md).

Тонкие обёртки над теми же сервисами и проверками прав, что и HTML-страницы
(core/views, core/services). Отличие от страниц — ответы: вместо перенаправлений
на вход и сообщений на странице — JSON с кодом 401/403/400/404.
Действия на страницах (способ, явка, отметки) — в core/views/api.py.
"""

import json
from functools import wraps

from django.contrib.auth import authenticate, login, logout
from django.core.paginator import Paginator
from django.db.models import Case, CharField, Q, Value, When
from django.db.models.functions import Concat
from django.http import HttpResponse, JsonResponse
from django.middleware.csrf import get_token
from django.utils import timezone
from django.views.decorators.csrf import ensure_csrf_cookie
from django.views.decorators.http import require_POST

from core.models import METHODS, UVZ, Employee
from core.services.excel import COLUMNS
from core.services.importing import (
    import_base,
    import_custom_report_archive,
    import_responsible_marks,
    import_turnout_hq_archive,
    import_voting_choices,
)
from core.services.reports.standard import (
    custom_reports_archive,
    reports_archive,
    responsible_marks_archive,
)
from core.views import exports
from core.views.common import (
    PER_PAGE,
    can_edit,
    filter_employees,
    filter_options,
    is_operator,
    is_uik_uvz_viewer,
    vote_counts,
)
from core.views.pages import CUSTOM_OKRUG_OPTIONS

# ФИО одной строкой: «Фамилия Имя» или «Фамилия Имя Отчество»
FIO = Concat(
    "surname",
    Value(" "),
    "name",
    Case(
        When(patronymic="", then=Value("")),
        default=Concat(Value(" "), "patronymic"),
    ),
    output_field=CharField(),
)


# --- Ответы и проверки ---


def json_response(data, status=200):
    return JsonResponse(
        data, status=status, safe=False, json_dumps_params={"ensure_ascii": False}
    )


def error(msg, status):
    return json_response({"status": "error", "msg": msg}, status)


def api_login_required(view):
    """Без входа — 401 с текстом, а не перенаправление на страницу входа."""

    @wraps(view)
    def wrapper(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return error("Необходима авторизация", 401)
        return view(request, *args, **kwargs)

    return wrapper


def operator_required(view):
    """Загрузки и шаблоны — только оператору (как на странице «Загрузка»)."""

    @wraps(view)
    def wrapper(request, *args, **kwargs):
        if not is_operator(request.user):
            return json_response({"error": "нет прав"}, 403)
        return view(request, *args, **kwargs)

    return api_login_required(wrapper)


def not_for_uvz_viewer(view):
    """Наблюдателю УИК-УВЗ доступна только явка: ответ «перейти на /elections»."""

    @wraps(view)
    def wrapper(request, *args, **kwargs):
        if is_uik_uvz_viewer(request.user):
            return json_response({"status": "redirect", "url": "/elections"}, 302)
        return view(request, *args, **kwargs)

    return api_login_required(wrapper)


def user_role(user):
    """Роль для фронтенда: operator, viewer, uik_uvz_viewer или guest."""
    if is_operator(user):
        return "operator"
    if not can_edit(user):
        return "viewer"
    if is_uik_uvz_viewer(user):
        return "uik_uvz_viewer"
    return "guest"


def paginated(rows, request):
    """Страница списка: неверный номер — первая страница, слишком большой — последняя."""
    paginator = Paginator(rows, PER_PAGE)
    page = paginator.get_page(request.GET.get("page"))
    return {
        "results": list(page.object_list),
        "pagination": {
            "current_page": page.number,
            "total_pages": paginator.num_pages,
            "total_count": paginator.count,
            "has_next": page.has_next(),
            "has_previous": page.has_previous(),
        },
    }


# --- Данные страниц ---


@not_for_uvz_viewer
def method_page(request):
    """Главная: сотрудники для выбора способа голосования."""
    rows = (
        filter_employees(request.GET)
        .annotate(fio=FIO)
        .values("department", "tab_number", "fio", "absence", "mark_uvz",
                "mark_deg", "okrug", "uik", "voted_method")  # fmt: skip
    )
    return json_response(paginated(rows, request))


@api_login_required
def elections_page(request):
    """Явка; наблюдатель УИК-УВЗ видит только выбравших или проголосовавших на УВЗ."""
    rows = filter_employees(request.GET).annotate(fio=FIO)
    fields = ["department", "tab_number", "fio", "okrug", "uik", "voted"]
    if is_uik_uvz_viewer(request.user):
        rows = rows.filter(Q(method=UVZ) | Q(voted_method=UVZ))
    else:
        fields.append("voted_method")
    return json_response(paginated(rows.values(*fields), request))


@not_for_uvz_viewer
def upload_page(request):
    """Страница загрузки: счётчики и колонки основного файла."""
    return json_response(
        {
            "status": "success",
            "counts": vote_counts(),
            "is_operator": is_operator(request.user),
            "columns": list(COLUMNS),
            "msg": request.session.pop("msg", ""),
        }
    )


@not_for_uvz_viewer
def export_page(request):
    """Страница отчётов: счётчики и списки для конструктора."""
    options = filter_options()
    return json_response(
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
            "uvz_not_voted_count": Employee.objects.filter(
                method=UVZ, mark_uvz=True, voted=False
            ).count(),
        }
    )


# --- Вход и выход ---


@ensure_csrf_cookie
def csrf(request):
    """CSRF-токен для POST-запросов фронтенда (заодно ставит cookie csrftoken)."""
    return JsonResponse({"csrfToken": get_token(request)})


@require_POST
def login_view(request):
    """Вход по логину и паролю из JSON {"username", "password"}."""
    if request.user.is_authenticated:
        return json_response(
            {"status": "success", "msg": "Пользователь уже авторизован"}
        )
    try:
        data = json.loads(request.body)
        username, password = data.get("username"), data.get("password")
    except (ValueError, AttributeError):
        return error("Невалидный JSON в теле запроса", 400)

    user = authenticate(request, username=username, password=password)
    if user is None:
        return error("Неверный логин или пароль", 400)
    login(request, user)
    return json_response(
        {
            "status": "success",
            "msg": "Вход выполнен успешно",
            "user": {
                "username": user.username,
                "role": user_role(user),
                "is_superuser": user.is_superuser,
            },
        }
    )


@require_POST
@api_login_required
def logout_view(request):
    logout(request)
    return json_response({"status": "success", "msg": "Сессия успешно завершена"})


# --- Загрузка файлов ---


def _upload(request, importer, extensions, message, details):
    """Общая загрузка: проверка расширения, импорт, ответ с итогами.

    message и details получают результат импортера (кортеж чисел).
    """
    upload = request.FILES.get("file")
    if not upload or not upload.name.lower().endswith(extensions):
        formats = " или ".join(extensions)
        return error(f"Ошибка: принимается только формат {formats}", 400)
    try:
        result = importer(upload)
    except ValueError as exc:
        return error(str(exc), 400)
    return json_response(
        {"status": "success", "msg": message(*result), "details": details(*result)}
    )


@require_POST
@operator_required
def upload_base(request):
    return _upload(
        request,
        import_base,
        (".xlsx",),
        lambda created, updated, total: (
            f"Строк в файле: {total}. Новых: {created}, изменено: {updated}, "
            f"без изменений: {total - created - updated}"
        ),
        lambda created, updated, total: {
            "total": total, "created": created, "updated": updated
        },
    )  # fmt: skip


@require_POST
@operator_required
def upload_voting_choices(request):
    return _upload(
        request,
        import_voting_choices,
        (".xlsx",),
        lambda updated, total, errors: (
            f"Обработано строк: {total}. Обновлено способов: {updated}, "
            f"ошибок/пропусков: {errors}"
        ),
        lambda updated, total, errors: {
            "total": total, "updated": updated, "errors_count": errors
        },
    )  # fmt: skip


def _changed_details(changed, total, errors):
    return {"total": total, "changed": changed, "errors_count": errors}


@require_POST
@operator_required
def upload_turnout_hq(request):
    return _upload(
        request,
        import_turnout_hq_archive,
        (".zip",),
        lambda changed, total, errors: (
            f"Обработано файлов в архиве. Всего строк {total}. "
            f"Успешно отмечено явок: {changed}, ошибок/пропусков: {errors}"
        ),
        _changed_details,
    )


@require_POST
@operator_required
def upload_custom_report(request):
    return _upload(
        request,
        import_custom_report_archive,
        (".zip", ".xlsx"),
        lambda changed, total, errors: (
            f"Обработано строк: {total}. Успешно обновлено записей: {changed}. "
            f"Ошибок/пропусков (нет в базе): {errors}."
        ),
        _changed_details,
    )


@require_POST
@operator_required
def upload_responsible_marks(request):
    return _upload(
        request,
        import_responsible_marks,
        (".zip", ".xlsx"),
        lambda changed, total, errors: (
            f"Обработано строк: {total}. Отмечено голосований через "
            f"ответственного: {changed}. Ошибок/пропусков: {errors}."
        ),
        _changed_details,
    )


# --- Выгрузки ---

# Готовые Excel-отчёты — те же, что на странице «Отчёты»
export_summary = api_login_required(exports.export_summary)
export_summary_no_19 = api_login_required(exports.export_summary_no_19)
export_productions = api_login_required(exports.export_productions)
export_production_methods = api_login_required(exports.export_production_methods)
export_production_methods_no_19 = api_login_required(
    exports.export_production_methods_no_19
)
export_employees = api_login_required(exports.export_employees)
export_uvz_statement_not_voted = api_login_required(
    exports.export_uvz_statement_not_voted
)
export_custom = api_login_required(exports.export_custom_report)


def _zip(archiver, filename, empty_msg):
    """ZIP-архив для скачивания или 404, если в архиве нет ни одного файла."""
    if not archiver.file_count:
        return error(empty_msg, 404)
    response = HttpResponse(archiver.build_bytes(), content_type="application/zip")
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _stamp():
    return f"{timezone.localtime():%Y%m%d_%H%M}"


@operator_required
def export_responsible_template(request):
    return _zip(
        responsible_marks_archive(),
        f"shablon_otvetstvenny_{_stamp()}.zip",
        "Нет данных для формирования архива",
    )


@api_login_required
def export_archive(request):
    return _zip(
        reports_archive(timezone.localtime(), "turnout"),
        f"yavka_po_ceham_{_stamp()}.zip",
        "Нет ни одного отдела — архив пустой",
    )


@api_login_required
def export_method_archive(request):
    return _zip(
        reports_archive(timezone.localtime(), "method"),
        f"sposob_po_ceham_{_stamp()}.zip",
        "Нет ни одного отдела — архив пустой",
    )


@api_login_required
def export_custom_archive(request):
    return _zip(
        custom_reports_archive(request.GET),
        f"svodny_po_ceham_{_stamp()}.zip",
        "Нет ни одного цеха — архив пустой",
    )
