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
})();
