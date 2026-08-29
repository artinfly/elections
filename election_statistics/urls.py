from django.urls import path

from . import views

urlpatterns = [
    # Главная страница — список сотрудников для простановки способа голосования
    path("", views.method_page, name="method"),
    # Страница входа в систему
    path("login/", views.login_view, name="login"),
    # Выход из системы
    path("logout/", views.logout_view, name="logout"),
    # Страница для отметки явки сотрудников
    path("elections/", views.elections_page, name="elections"),
    # Страница загрузки базы сотрудников из Excel (только для операторов)
    path("upload/", views.upload_page, name="upload"),
    # Обработчик POST-запроса с файлом для импорта базы
    path("upload/base/", views.upload_base, name="upload_base"),
    # Страница со списком всех доступных отчётов и формой конструктора сводного отчёта
    path("export/", views.export_page, name="export"),
    # Выгрузка полной таблицы со всеми сотрудниками
    path("export/employees/", views.export_employees, name="export_employees"),
    # Сводная таблица по цехам (все способы + явка)
    path("export/summary/", views.export_summary, name="export_summary"),
    # Сводная таблица по цехам, но без сотрудников 19 округа
    path(
        "export/summary-no-u19/",
        views.export_summary_no_19,
        name="export_summary_no_19",
    ),
    # Отчёт по производствам: сколько человек работает и сколько проголосовало
    path(
        "export/productions/",
        views.export_productions,
        name="export_productions",
    ),
    # Отчёт по производствам: распределение способов голосования (ДЭГ, УИК, УВЗ, У19)
    path(
        "export/productions-methods/",
        views.export_production_methods,
        name="export_production_methods",
    ),
    # Отчёт по производствам: способы голосования, но без сотрудников 19 округа
    path(
        "export/productions-methods-no-u19/",
        views.export_production_methods_no_19,
        name="export_production_methods_no_19",
    ),
    # ZIP-архив с отдельными отчётами по явке для каждого цеха
    path("export/archive/", views.export_archive, name="export_archive"),
    # ZIP-архив с отдельными отчётами по способам голосования для каждого цеха
    path(
        "export/archive-methods/",
        views.export_method_archive,
        name="export_method_archive",
    ),
    # Сводный отчёт по сотрудникам на основе фильтров из формы (УИК, цех, таб.№, ФИО, округ и отметки по способам)
    path(
        "export/custom/",
        views.export_custom_report,
        name="export_custom",
    ),
    # API: обновление способа голосования для одного сотрудника
    path("api/method/", views.api_method, name="api_method"),
    # API: отметка явки для одного сотрудника
    path("api/voted/", views.api_voted, name="api_voted"),
    # API: массовая отметка явки для всех сотрудников, попадающих под фильтры
    path("api/bulk-voted/", views.api_bulk_voted, name="api_bulk_voted"),
    # API: статистика по УИКам (используется в модальном окне на страницах method/elections)
    path("api/uik-stats/", views.api_uik_stats, name="api_uik_stats"),
]
