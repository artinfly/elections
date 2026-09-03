/**
 * Главный скрипт приложения.
 *
 * Описание:
 *   Содержит общую логику для всех страниц:
 *   1. Чтение CSRF-токена для безопасных POST-запросов.
 *   2. Функцию postJSON для асинхронного взаимодействия с API.
 *   3. Логику модального окна со статистикой по УИК.
 *
 *   Обернут в IIFE (Immediately Invoked Function Expression),
 *   чтобы не засорять глобальную область видимости.
 */
(function () {
    "use strict";

    // ========================================================================
    // CSRF-защита
    // ========================================================================
    // Токен читается из скрытой формы в base.html и подставляется в заголовок
    // X-CSRFToken для всех POST-запросов. Это стандартный механизм защиты Django.
    var field = document.querySelector('input[name="csrfmiddlewaretoken"]');
    var token = field ? field.value : "";

    // ========================================================================
    // Утилиты
    // ========================================================================

    /**
     * Показывает сообщение об ошибке в общем контейнере #app-error.
     *
     * @param {string} message - Текст ошибки.
     */
    function showError(message) {
        var box = document.getElementById("app-error");
        if (!box) return;
        box.textContent = message;
        box.style.display = "block";
    }

    /**
     * Экранирует HTML-сущности для защиты от XSS при выводе данных из API.
     *
     * @param {any} value - Значение для экранирования.
     * @returns {string} Безопасная строка.
     */
    function escapeHtml(value) {
        return String(value === null || value === undefined ? "" : value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
    }

    // ========================================================================
    // Работа с фильтрами и API
    // ========================================================================

    /**
     * Собирает текущие фильтры из адресной строки браузера.
     *
     * Описание:
     *   Парсит GET-параметры URL, исключая параметр пагинации 'page'.
     *   Используется для того, чтобы действия (отметка явки, смена способа)
     *   применялись только к текущей выборке, а не ко всей базе.
     *
     * @returns {Object} Словарь фильтров {ключ: значение}.
     */
    window.currentFilters = function () {
        var filters = {};
        new URLSearchParams(window.location.search).forEach(function (value, key) {
            if (key !== "page") filters[key] = value;
        });
        return filters;
    };

    /**
     * Выполняет POST-запрос к API с JSON-телом и CSRF-заголовком.
     *
     * Описание:
     *   Универсальная функция для всех асинхронных действий.
     *   Автоматически добавляет текущие фильтры в тело запроса, если они
     *   не были переданы явно. При ошибке показывает сообщение пользователю.
     *
     * @param {string} url - URL эндпоинта API.
     * @param {Object} data - Данные для отправки.
     * @returns {Promise<Object>} Промис с JSON-ответом сервера.
     */
    window.postJSON = function (url, data) {
        // Если фильтры не переданы явно, подтягиваем их из URL.
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
                showError("Не удалось сохранить: " + error.message);
                throw error;
            });
    };

    // ========================================================================
    // Модальное окно "Люди по участкам"
    // ========================================================================
    // Логика ниже выполняется только если на странице есть карточка УИК и модалка.
    var modal = document.getElementById("uik-modal");
    var card = document.getElementById("uik-card");
    if (!modal || !card) return;

    var body = document.getElementById("uik-modal-body");
    var sub = document.getElementById("uik-modal-sub");
    
    // Флаг включает колонки явки в таблице модалки (активно на странице elections).
    var withTurnout = card.dataset.turnout === "1";

    /**
     * Скрывает модальное окно.
     */
    function closeModal() {
        modal.hidden = true;
    }

    /**
     * Отрисовывает таблицу участков внутри модального окна.
     *
     * @param {Array<Object>} rows - Массив объектов {uik, people, came} из API.
     */
    function renderModal(rows) {
        if (!rows.length) {
            body.innerHTML = '<div class="modal-empty">Нет данных</div>';
            sub.textContent = "";
            return;
        }
        
        var totalPeople = 0;
        var totalCame = 0;
        
        // Генерация строк таблицы.
        var cells = rows.map(function (r) {
            totalPeople += r.people;
            totalCame += r.came;
            var share = r.people ? Math.round((r.came / r.people) * 100) : 0;
            
            var row = "<tr><td>" + (r.uik ? escapeHtml(r.uik) : "<i>не указан</i>") +
                      "</td><td>" + r.people + "</td>";
                      
            if (withTurnout) {
                row += "<td>" + r.came + "</td><td>" + share + "%</td>";
            }
            return row + "</tr>";
        }).join("");

        // Генерация заголовков и итоговой строки.
        var head = "<th>УИК</th><th>Человек</th>";
        var foot = '<tr class="modal-total"><td>Итого</td><td>' + totalPeople + "</td>";
        
        if (withTurnout) {
            head += "<th>Проголосовало</th><th>Явка</th>";
            var totalShare = totalPeople ? Math.round((totalCame / totalPeople) * 100) : 0;
            foot += "<td>" + totalCame + "</td><td>" + totalShare + "%</td>";
        }
        foot += "</tr>";

        body.innerHTML = '<table class="modal-table"><thead><tr>' + head +
                         "</tr></thead><tbody>" + cells + foot + "</tbody></table>";
                         
        sub.textContent = "участков: " + rows.length + ", человек: " + totalPeople;
    }

    // Обработчик открытия модалки по клику на карточку УИК.
    card.addEventListener("click", function () {
        modal.hidden = false;
        body.innerHTML = '<div class="modal-empty">Загрузка...</div>';
        sub.textContent = "";
        
        // Формируем строку запроса из текущих фильтров.
        var query = new URLSearchParams(window.currentFilters()).toString();
        
        fetch("/api/uik-stats/" + (query ? "?" + query : ""))
            .then(function (r) {
                if (!r.ok) throw new Error("сервер ответил " + r.status);
                return r.json();
            })
            .then(renderModal)
            .catch(function (error) {
                body.innerHTML = '<div class="modal-empty">Не удалось загрузить: ' +
                                 escapeHtml(error.message) + "</div>";
            });
    });

    // Обработчики закрытия модалки.
    document.getElementById("uik-modal-close").addEventListener("click", closeModal);
    
    // Закрытие по клику на затемненный фон.
    modal.addEventListener("click", function (e) {
        if (e.target === modal) closeModal();
    });
    
    // Закрытие по нажатию клавиши Escape.
    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && !modal.hidden) closeModal();
    });
})();