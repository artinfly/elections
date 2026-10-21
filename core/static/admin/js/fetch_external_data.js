/**
 * Админка пользователей: кнопка «Получить данные сотрудника».
 *
 * Берёт табельный номер из поля username, запрашивает данные у сервера
 * (view fetch_external_data в admin.py обращается к кадровому API)
 * и заполняет поля формы: ФИО, e-mail, отчество, API-ключ, «уволен».
 * Адрес запроса задан в data-url-template кнопки, вместо EMP_PLACEHOLDER
 * подставляется табельный номер.
 */
document.addEventListener('DOMContentLoaded', function () {
    var btn = document.getElementById('fetch-external-data-btn');
    if (!btn) return;  // кнопка есть только в формах пользователя

    btn.addEventListener('click', function () {
        var usernameField = document.getElementById('id_username');
        var status = document.getElementById('fetch-external-data-status');

        if (!usernameField) {
            status.textContent = 'Не найдено поле username';
            return;
        }

        // Логин пользователя — это его табельный номер
        var empNumber = usernameField.value.trim();

        if (!empNumber) {
            status.textContent = 'Сначала заполните username';
            return;
        }

        var url = btn.dataset.urlTemplate.replace('EMP_PLACEHOLDER', encodeURIComponent(empNumber));

        status.textContent = 'Загрузка...';

        fetch(url, {
            method: 'GET',
            headers: { 'X-Requested-With': 'XMLHttpRequest' },
            credentials: 'same-origin'
        })
        .then(function (response) {
            // Сервер при ошибке отвечает JSON {"error": "..."}
            if (!response.ok) {
                return response.json().then(function (data) {
                    throw new Error(data.error || 'Ошибка запроса');
                });
            }
            return response.json();
        })
        .then(function (data) {
            fillField('id_first_name', data.name);
            fillField('id_last_name', data.surname);
            fillField('id_email', data.email);
            fillField('id_patronymic', data.patronymic);
            fillField('id_api_key', data.api_key);

            // «Уволен» — флажок, у него меняется checked, а не value
            var isFiredField = document.getElementById('id_is_fired');
            if (isFiredField) {
                isFiredField.checked = Boolean(data.is_fired);
            }

            status.textContent = 'Готово';
        })
        .catch(function (err) {
            status.textContent = 'Ошибка: ' + err.message;
        });
    });

    /**
     * Записывает значение в поле формы, если такое поле есть на странице
     * и значение пришло в ответе.
     */
    function fillField(id, value) {
        var el = document.getElementById(id);
        if (el && value !== undefined) {
            el.value = value;
        }
    }
});
