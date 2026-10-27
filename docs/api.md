# JSON API для фронтенда

Приложение `apps/api`: те же данные и действия, что на страницах сайта, но в формате JSON —
для отдельного фронтенда. Права те же, что на страницах.

## Вход и CSRF

Вход — сессией Django (cookie `sessionid`). Все POST-запросы требуют CSRF-токен
в заголовке `X-CSRFToken`.

1. `GET /api/csrf/` → `{"csrfToken": "…"}` (заодно ставит cookie `csrftoken`).
2. `POST /api/login/` с телом `{"username": "…", "password": "…"}`.
3. Дальше запросы идут с cookie (во `fetch` — `credentials: "include"`).
4. `POST /api/logout/` — выход.

Ответ входа:

```json
{"status": "success", "msg": "Вход выполнен успешно",
 "user": {"username": "ivanov", "role": "operator", "is_superuser": false}}
```

`role`: `operator` (оператор или суперпользователь), `viewer` (только просмотр),
`uik_uvz_viewer` (наблюдатель УИК-УВЗ — видит только явку), `guest` (без роли).

Если фронтенд открывается с другого адреса, этот адрес указывается в `.env`:
`CORS_ALLOWED_ORIGINS=http://localhost:7001` (нужна библиотека `django-cors-headers`).

## Ошибки

| Код | Когда | Тело |
|---|---|---|
| 400 | неверный файл, JSON, логин или пароль | `{"status": "error", "msg": "…"}` |
| 401 | не выполнен вход | `{"status": "error", "msg": "Необходима авторизация"}` |
| 403 | загрузка или шаблон без роли оператора | `{"error": "нет прав"}` |
| 404 | архив пустой | `{"status": "error", "msg": "…"}` |
| 302 | страница недоступна наблюдателю УИК-УВЗ | `{"status": "redirect", "url": "/elections"}` |
| 405 | не тот метод (GET вместо POST) | — |

## Данные страниц (GET)

Фильтры — параметры адреса, как на страницах: `q` (ФИО или табельный номер), `dep`, `production`,
`service`, `uik`, `okrug` (`none` — пусто), `method` (`deg`, `uik`, `uvz`, `u19`, `none`),
`where` (фактический способ), `voted` (`yes`/`no`), `page`.

| Адрес | Что возвращает |
|---|---|
| `/api/` | сотрудники для выбора способа: `department`, `tab_number`, `fio`, `absence`, `mark_uvz`, `mark_deg`, `okrug`, `uik`, `voted_method` |
| `/api/elections/` | сотрудники для отметки явки: `department`, `tab_number`, `fio`, `okrug`, `uik`, `voted`, `voted_method` (наблюдателю УИК-УВЗ — только выбравшие или проголосовавшие на УВЗ, без `voted_method`) |
| `/api/upload/` | `counts` (счётчики), `is_operator`, `columns` (колонки основного файла), `msg` |
| `/api/export/` | `counts`, `is_operator`, списки для конструктора (`departments`, `productions`, `services`, `uiks`, `methods`, `custom_okrugs`), `uvz_not_voted_count` |

Списки — по 100 строк:

```json
{"results": [...],
 "pagination": {"current_page": 1, "total_pages": 5, "total_count": 431,
                "has_next": true, "has_previous": false}}
```

Неверный номер страницы — первая страница, слишком большой — последняя.

`counts`:

```json
{"method": {"deg": 0, "uik": 0, "uvz": 0, "u19": 0, "none": 0, "total": 0},
 "voted":  {"deg": 0, "uik": 0, "uvz": 0, "u19": 0, "none": 0, "total": 0},
 "no_uik": 0}
```

## Загрузка файлов (POST, роль оператора)

Файл — в поле `file` (multipart/form-data). Ответ:
`{"status": "success", "msg": "итог текстом", "details": {...}}`.

| Адрес | Файл | `details` |
|---|---|---|
| `/api/upload/base/` | .xlsx — основной список сотрудников | `total`, `created`, `updated` |
| `/api/upload/voting-choices/` | .xlsx — способы голосования из отчёта штаба («1» в колонке способа) | `total`, `updated`, `errors_count` |
| `/api/upload/turnout-hq/` | .zip — отчёты штаба с явкой | `total`, `changed`, `errors_count` |
| `/api/upload/custom-report/` | .zip или .xlsx — сводный отчёт | `total`, `changed`, `errors_count` |
| `/api/upload/responsible/` | .zip или .xlsx — голосование через ответственного | `total`, `changed`, `errors_count` |

## Выгрузки (GET, файл для скачивания)

| Адрес | Файл |
|---|---|
| `/api/export/summary/` | сводная по цехам (.xlsx) |
| `/api/export/summary-no-u19/` | сводная по цехам без 19 округа |
| `/api/export/productions/` | по производствам |
| `/api/export/productions-methods/` | способы по производствам |
| `/api/export/productions-methods-no-u19/` | способы по производствам без 19 округа |
| `/api/export/employees/` | все сотрудники |
| `/api/export/uvz-statement-not-voted/` | заявление на УИК-УВЗ есть, но не проголосовал |
| `/api/export/custom/` | конструктор отчёта (параметры как на странице «Отчёты», `grouping`) |
| `/api/export/archive/` | ZIP: явка по цехам |
| `/api/export/archive-methods/` | ZIP: способы по цехам |
| `/api/export/custom-archive/` | ZIP: конструктор по цехам |
| `/api/export/responsible-template/` | ZIP: шаблоны «через ответственного» (роль оператора) |

## Действия на страницах (POST, JSON)

Используются и страницами сайта (`core/views/api.py`):
`/api/method/` (способ), `/api/mark/` (отметка регистрации ДЭГ/УВЗ), `/api/voted/` (явка),
`/api/bulk-voted/` (явка по фильтру), `/api/toggle-absence/<id>/` (отсутствие по УП),
`GET /api/uik-stats/` (статистика по УИКам).
