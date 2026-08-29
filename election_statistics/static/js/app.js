(function () {
    "use strict";

    // CSRF-токен читается из скрытой формы #csrf-holder в base.html
    // и подставляется в заголовок X-CSRFToken у всех POST-запросов к API
    var field = document.querySelector('input[name="csrfmiddlewaretoken"]');
    var token = field ? field.value : "";

    // Показывает сообщение в общем блоке ошибок #app-error (под шапкой страницы)
    function show(message) {
        var box = document.getElementById("app-error");
        if (!box) return;
        box.textContent = message;
        box.style.display = "block";
    }

    // Собирает текущие фильтры из адресной строки (все GET-параметры кроме page).
    // Используются, чтобы после изменения данных сервер пересчитал счётчики
    // по той же выборке, которую видит пользователь.
    window.currentFilters = function () {
        var filters = {};
        new URLSearchParams(window.location.search).forEach(function (value, key) {
            if (key !== "page") filters[key] = value;
        });
        return filters;
    };

    // Универсальный POST-запрос к API с JSON-телом и CSRF-заголовком.
    // Автоматически добавляет текущие фильтры, если они не переданы явно.
    // При ошибке показывает сообщение и пробрасывает исключение дальше.
    window.postJSON = function (url, data) {
        data.filters = data.filters || window.currentFilters();
        return fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "X-CSRFToken": token,
            },
            body: JSON.stringify(data),
        })
            .then(function (response) {
                if (!response.ok) {
                    throw new Error("сервер ответил " + response.status);
                }
                return response.json();
            })
            .catch(function (error) {
                show("Не удалось сохранить: " + error.message);
                throw error;
            });
    };

    // ---------- Модальное окно "Люди по участкам" ----------
    // Окно и карточка-кнопка есть только на страницах method/elections;
    // на остальных страницах скрипт дальше не работает.
    var modal = document.getElementById("uik-modal");
    var card = document.getElementById("uik-card");
    if (!modal || !card) return;

    var body = document.getElementById("uik-modal-body");
    var sub = document.getElementById("uik-modal-sub");
    // data-turnout="1" (страница elections) включает колонки явки в таблице модалки
    var withTurnout = card.dataset.turnout === "1";

    // Скрывает модальное окно
    function close() {
        modal.hidden = true;
    }

    // Экранирует HTML в данных из API, чтобы нельзя было внедрить разметку
    function escape(value) {
        return String(value === null || value === undefined ? "" : value)
            .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    // Рисует таблицу участков: УИК, человек (и явка с процентом, если withTurnout).
    // Внизу добавляет строку "Итого" с суммами и общим процентом.
    function render(rows) {
        if (!rows.length) {
            body.innerHTML = '<div class="modal-empty">Нет данных</div>';
            sub.textContent = "";
            return;
        }
        var people = 0;
        var came = 0;
        var cells = rows.map(function (r) {
            people += r.people;
            came += r.came;
            var share = r.people ? Math.round((r.came / r.people) * 100) : 0;
            var row = "<tr><td>" + (r.uik ? escape(r.uik) : "<i>не указан</i>") +
                "</td><td>" + r.people + "</td>";
            if (withTurnout) {
                row += "<td>" + r.came + "</td><td>" + share + "%</td>";
            }
            return row + "</tr>";
        }).join("");

        var head = "<th>УИК</th><th>Человек</th>";
        var foot = '<tr class="modal-total"><td>Итого</td><td>' + people + "</td>";
        if (withTurnout) {
            head += "<th>Проголосовало</th><th>Явка</th>";
            foot += "<td>" + came + "</td><td>" +
                (people ? Math.round((came / people) * 100) : 0) + "%</td>";
        }
        foot += "</tr>";

        body.innerHTML = '<table class="modal-table"><thead><tr>' + head +
            "</tr></thead><tbody>" + cells + foot + "</tbody></table>";
        sub.textContent = "участков: " + rows.length + ", человек: " + people;
    }

    // Клик по карточке УИК: открывает окно и тянет статистику из /api/uik-stats/
    // с теми же фильтрами, что применены на странице
    card.addEventListener("click", function () {
        modal.hidden = false;
        body.innerHTML = '<div class="modal-empty">Загрузка...</div>';
        sub.textContent = "";
        var query = new URLSearchParams(window.currentFilters()).toString();
        fetch("/api/uik-stats/" + (query ? "?" + query : ""))
            .then(function (r) {
                if (!r.ok) throw new Error("сервер ответил " + r.status);
                return r.json();
            })
            .then(render)
            .catch(function (error) {
                body.innerHTML = '<div class="modal-empty">Не удалось загрузить: ' +
                    escape(error.message) + "</div>";
            });
    });

    // Закрытие окна: кнопка-крестик, клик по фону вокруг окна, клавиша Escape
    document.getElementById("uik-modal-close").addEventListener("click", close);
    modal.addEventListener("click", function (e) {
        if (e.target === modal) close();
    });
    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && !modal.hidden) close();
    });
})();
