"""
Файл маршрутов (ссылок) приложения.

Описание:
    Здесь прописывается связь между адресом в браузере (например, /login/)
    и функцией в файле views.py, которая должна сработать при переходе по нему.
    Каждый маршрут имеет имя (параметр name), которое используется в коде
    и шаблонах для генерации ссылок и перенаправлений.
"""

from django.urls import path

from . import views

urlpatterns = [
    # ==============================================================================
    # Аутентификация (вход и выход из системы)
    # ==============================================================================
    # Страница входа с логином и паролем.
    path("login/", views.login_view, name="login"),
    # Завершение сеанса и возврат на страницу входа.
    path("logout/", views.logout_view, name="logout"),
    # ==============================================================================
    # Основные страницы интерфейса
    # ==============================================================================
    # Главная страница: выбор способа голосования (план).
    path("", views.method_page, name="method"),
    # Страница отметки фактической явки.
    path("elections/", views.elections_page, name="elections"),
    # Страница загрузки списков сотрудников из Excel.
    path("upload/", views.upload_page, name="upload"),
    # Страница выгрузки отчетов и формирования сводок.
    path("export/", views.export_page, name="export"),
    # ==============================================================================
    # Обработчики загрузки файлов (вызываются формами на странице /upload/)
    # ==============================================================================
    # Импорт основного списка сотрудников (ФИО, адреса, табельные номера).
    path("upload/base/", views.upload_base, name="upload_base"),
    # Импорт данных о выборе способа голосования из отчета штаба.
    path(
        "upload/voting-choices/",
        views.upload_voting_choices,
        name="upload_voting_choices",
    ),
    # ==============================================================================
    # Экспорт отчетов в Excel и ZIP (скачивание файлов)
    # ==============================================================================
    # Выгрузка полного списка всех сотрудников со всеми данными.
    path("export/employees/", views.export_employees, name="export_employees"),
    # Сводная таблица по цехам (план голосования и явка).
    path("export/summary/", views.export_summary, name="export_summary"),
    # Сводная таблица по цехам, исключая сотрудников 19-го округа.
    path(
        "export/summary-no-u19/",
        views.export_summary_no_19,
        name="export_summary_no_19",
    ),
    # Отчет по производствам (группировка по службам).
    path("export/productions/", views.export_productions, name="export_productions"),
    # Способы голосования по производствам.
    path(
        "export/productions-methods/",
        views.export_production_methods,
        name="export_production_methods",
    ),
    # Способы голосования по производствам без учета 19-го округа.
    path(
        "export/productions-methods-no-u19/",
        views.export_production_methods_no_19,
        name="export_production_methods_no_19",
    ),
    # ZIP-архив с отчетами по явке (отдельный файл на каждый цех).
    path("export/archive/", views.export_archive, name="export_archive"),
    # ZIP-архив с отчетами по выбору способа (отдельный файл на каждый цех).
    path(
        "export/archive-methods/",
        views.export_method_archive,
        name="export_method_archive",
    ),
    # Кастомный сводный отчет (конструктор фильтров).
    path("export/custom/", views.export_custom_report, name="export_custom"),
    # Кастомный сводный отчет в виде ZIP-архива по цехам.
    path(
        "export/custom-archive/",
        views.export_custom_archive,
        name="export_custom_archive",
    ),
    # ==============================================================================
    # API (асинхронные запросы без перезагрузки страницы)
    # ==============================================================================
    # Обновление запланированного способа голосования для одного сотрудника.
    path("api/method/", views.api_method, name="api_method"),
    # Простановка отметок о регистрации (ДЭГ или УВЗ).
    path("api/mark/", views.api_mark, name="api_mark"),
    # Отметка явки для одного сотрудника (проголосовал/не проголосовал).
    path("api/voted/", views.api_voted, name="api_voted"),
    # Массовая отметка явки по текущему фильтру.
    path("api/bulk-voted/", views.api_bulk_voted, name="api_bulk_voted"),
    # Получение статистики по УИКам для модального окна.
    path("api/uik-stats/", views.api_uik_stats, name="api_uik_stats"),
    # Переключение отметки «Отсутствие по УП» для кнопки на странице способа.
    # Путь изменен на api/ для соответствия общим правилам проекта.
    path(
        "api/toggle-absence/<int:employee_id>/",
        views.api_toggle_absence,
        name="toggle_absence",
    ),
]
