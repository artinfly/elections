(function () {
    "use strict";

    var field = document.querySelector('input[name="csrfmiddlewaretoken"]');
    var token = field ? field.value : "";

    function show(message) {
        var box = document.getElementById("app-error");
        if (!box) return;
        box.textContent = message;
        box.style.display = "block";
    }

    window.currentFilters = function () {
        var filters = {};
        new URLSearchParams(window.location.search).forEach(function (value, key) {
            if (key !== "page") filters[key] = value;
        });
        return filters;
    };

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

    var modal = document.getElementById("uik-modal");
    var card = document.getElementById("uik-card");
    if (!modal || !card) return;

    var body = document.getElementById("uik-modal-body");
    var sub = document.getElementById("uik-modal-sub");
    var withTurnout = card.dataset.turnout === "1";

    function close() {
        modal.hidden = true;
    }

    function escape(value) {
        return String(value === null || value === undefined ? "" : value)
            .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

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

    document.getElementById("uik-modal-close").addEventListener("click", close);
    modal.addEventListener("click", function (e) {
        if (e.target === modal) close();
    });
    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && !modal.hidden) close();
    });
})();
