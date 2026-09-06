document.addEventListener('DOMContentLoaded', function () {
    var btn = document.getElementById('fetch-external-data-btn');
    if (!btn) return;

    btn.addEventListener('click', function () {
        var usernameField = document.getElementById('id_username');
        var status = document.getElementById('fetch-external-data-status');

        if (!usernameField) {
            status.textContent = 'Не найдено поле username';
            return;
        }

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

    function fillField(id, value) {
        var el = document.getElementById(id);
        if (el && value !== undefined) {
            el.value = value;
        }
    }
});