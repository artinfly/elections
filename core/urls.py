"""
Файл маршрутов (ссылок) приложения.

Описание:
    Здесь прописывается связь между адресом в браузере (например, /login/)
    и функцией в пакете views/, которая должна сработать при переходе.
    Каждый маршрут имеет имя (параметр name), которое используется в коде
    и шаблонах для генерации ссылок и перенаправлений.
"""

from django.urls import path

from .views import api, auth, exports, pages, uploads

urlpatterns = [
    # ==============================================================================
    # Аутентификация (вход и выход из системы)
    # ==============================================================================
    # Страница входа с логином и паролем.
    path("login/", auth.login_view, name="login"),
    # Завершение сеанса и возврат на страницу входа.
    path("logout/", auth.logout_view, name="logout"),
    # ==============================================================================
    # Основные страницы интерфейса
    # ==============================================================================
    # Главная страница: выбор способа голосования (план).
    path("", pages.method_page, name="method"),
    # Страница отметки фактической явки.
    path("elections/", pages.elections_page, name="elections"),
    # Страница загрузки списков сотрудников из Excel.
    path("upload/", pages.upload_page, name="upload"),
    # Страница выгрузки отчетов и формирования сводок.
    path("export/", pages.export_page, name="export"),
    # ==============================================================================
    # Обработчики загрузки файлов (вызываются формами на странице /upload/)
    # ==============================================================================
    # Импорт основного списка сотрудников (ФИО, адреса, табельные номера).
    path("upload/base/", uploads.upload_base, name="upload_base"),
    # Импорт явки из ZIP-архива отчётов штаба с датой и временем голосования.
    path("upload/turnout-hq/", uploads.upload_turnout_hq, name="upload_turnout_hq"),
    # Импорт данных из кастомного сводного отчёта (Excel или ZIP с Excel).
    path(
        "upload/custom-report/",
        uploads.upload_custom_report,
        name="upload_custom_report",
    ),
    # Импорт отметок «Голосование через ответственного» из Excel или ZIP.
    path(
        "upload/responsible/",
        uploads.upload_responsible_marks,
        name="upload_responsible_marks",
    ),
    # ==============================================================================
    # Экспорт отчетов в Excel и ZIP (скачивание файлов)
    # ==============================================================================
    # Выгрузка полного списка всех сотрудников со всеми данными.
    path("export/employees/", exports.export_employees, name="export_employees"),
    # Список: заявление на УИК-УВЗ оформлено, но не проголосовал.
    path(
        "export/uvz-statement-not-voted/",
        exports.export_uvz_statement_not_voted,
        name="export_uvz_statement_not_voted",
    ),
    # Шаблон Excel для заполнения отметок «Голосование через ответственного».
    path(
        "export/responsible-template/",
        exports.export_responsible_template,
        name="export_responsible_template",
    ),
    # Сводная таблица по цехам (план голосования и явка).
    path("export/summary/", exports.export_summary, name="export_summary"),
    # Сводная таблица по цехам, исключая сотрудников 19-го округа.
    path(
        "export/summary-no-u19/",
        exports.export_summary_no_19,
        name="export_summary_no_19",
    ),
    # Отчет по производствам (группировка по службам).
    path("export/productions/", exports.export_productions, name="export_productions"),
    # Способы голосования по производствам.
    path(
        "export/productions-methods/",
        exports.export_production_methods,
        name="export_production_methods",
    ),
    # Способы голосования по производствам без учета 19 округа.
    path(
        "export/productions-methods-no-u19/",
        exports.export_production_methods_no_19,
        name="export_production_methods_no_19",
    ),
    # ZIP-архив с отчетами по явке (отдельный файл на каждый цех).
    path("export/archive/", exports.export_archive, name="export_archive"),
    # ZIP-архив с отчетами по выбору способа (отдельный файл на каждый цех).
    path(
        "export/archive-methods/",
        exports.export_method_archive,
        name="export_method_archive",
    ),
    # Кастомный сводный отчет (конструктор фильтров).
    path("export/custom/", exports.export_custom_report, name="export_custom"),
    # Кастомный сводный отчет в виде ZIP-архива по цехам.
    path(
        "export/custom-archive/",
        exports.export_custom_archive,
        name="export_custom_archive",
    ),
    # ==============================================================================
    # API (асинхронные запросы без перезагрузки страницы)
    # ==============================================================================
    # Обновление запланированного способа голосования для одного сотрудника.
    path("api/method/", api.api_method, name="api_method"),
    # Простановка отметок о регистрации (ДЭГ или УВЗ).
    path("api/mark/", api.api_mark, name="api_mark"),
    # Отметка явки для одного сотрудника (проголосовал/не проголосовал).
    path("api/voted/", api.api_voted, name="api_voted"),
    # Массовая отметка явки по текущему фильтру.
    path("api/bulk-voted/", api.api_bulk_voted, name="api_bulk_voted"),
    # Получение статистики по УИКам для модального окна.
    path("api/uik-stats/", api.api_uik_stats, name="api_uik_stats"),
    # Переключение отметки «Отсутствие по УП» для конкретного сотрудника.
    path(
        "api/toggle-absence/<int:employee_id>/",
        api.api_toggle_absence,
        name="toggle_absence",
    ),
]
