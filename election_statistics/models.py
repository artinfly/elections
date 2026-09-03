"""
Модели приложения election_statistics.

Единственная таблица — Employee: личные данные сотрудника, адрес,
привязка к УИК и округу, план и факт голосования.
Здесь же — коды способов голосования, которые используют views,
отчёты, импортеры и шаблоны.
"""

from django.db import models

# Коды способов голосования — короткие значения, которые хранятся в базе
DEG = "deg"
UIK = "uik"
UVZ = "uvz"
UIK19 = "u19"

# Пары (код, подпись) для choices в модели и для выпадающих списков в шаблонах
METHODS = [(DEG, "ДЭГ"), (UIK, "УИК"), (UVZ, "УИК-УВЗ"), (UIK19, "УИК-19")]

# Словарь код -> человекочитаемая подпись для отображения в таблицах и отчётах
METHOD_LABELS = dict(METHODS)


class Employee(models.Model):
    """
    Сотрудник — единственная таблица базы.

    Хранит личные данные, адрес, привязку к УИК и округу,
    запланированный способ голосования и отметки явки.
    """

    # Табельный номер — уникальный ключ для импорта и поиска
    tab_number = models.CharField("Таб№", max_length=20, unique=True)
    # Цех (подразделение) — основная группировка в отчётах по цехам
    department = models.CharField("Подразделение", max_length=50, db_index=True)
    # Производство из внешней базы (import_workshops);
    # используется как фильтр на страницах. Внимание: отчёты
    # «по производствам» группируются по полю service — см. комментарий
    # в reports.py, не «исправлять» без синхронного переименования
    production = models.CharField(
        "Производство", max_length=200, blank=True, db_index=True
    )
    # Служба; именно по этому полю группируются отчёты «по производствам»
    service = models.CharField("Служба", max_length=200, blank=True, db_index=True)
    surname = models.CharField("Фамилия", max_length=100)
    name = models.CharField("Имя", max_length=100)
    patronymic = models.CharField("Отчество", max_length=100, blank=True)
    position = models.CharField("Должность", max_length=255, blank=True)
    category = models.CharField("Категория", max_length=100, blank=True)
    birth_date = models.DateField("Дата рождения", null=True, blank=True)
    region = models.CharField("Регион", max_length=100, blank=True)
    city = models.CharField("Город", max_length=100, blank=True)
    street = models.CharField("Улица", max_length=150, blank=True)
    house = models.CharField("Дом", max_length=20, blank=True)
    # Номер участка, к которому приписан сотрудник
    uik = models.CharField("УИК", max_length=10, blank=True, db_index=True)
    uik_address = models.CharField("Адрес УИК", max_length=255, blank=True)
    district = models.CharField("Район", max_length=100, blank=True)
    # Избирательный округ (19, 20, 21 или пусто)
    okrug = models.CharField("Округ", max_length=10, blank=True)
    # Запланированный способ голосования (что выбрал сотрудник)
    method = models.CharField(
        "Способ голосования", max_length=3, choices=METHODS, blank=True, db_index=True
    )
    # Отметка явки: проголосовал ли сотрудник
    voted = models.BooleanField("Проголосал", default=False, db_index=True)
    # Способ, которым сотрудник фактически проголосовал (заполняется при voted=True)
    voted_method = models.CharField(
        "Где голосовал", max_length=3, choices=METHODS, blank=True, db_index=True
    )
    # Время, когда была проставлена отметка явки
    voted_at = models.DateTimeField("Время отметки", null=True, blank=True)
    # Отметка "открепился": сотрудник будет голосовать вне своего округа
    detached = models.BooleanField("Открепился", default=False)
    # Отметка "не пойдет": сотрудник заявил, что не будет участвовать в голосовании
    not_going = models.BooleanField("Не пойдет", default=False)
    mark_uvz = models.BooleanField("Регистрация на УИК-УВЗ", default=False)
    mark_deg = models.BooleanField("Регистрация на ДЭГ", default=False)
    absence = models.BooleanField("Отсутствие по УП", default=False)

    class Meta:
        verbose_name = "Сотрудник"
        verbose_name_plural = "Сотрудники"
        # Сортировка по умолчанию — по ФИО.
        # Важно: в агрегирующих запросах (.values().annotate()) эту сортировку
        # нужно сбрасывать пустым .order_by(), иначе поля сортировки попадут
        # в GROUP BY и разобьют группы. В reports.py сброс уже стоит.
        ordering = ["surname", "name", "patronymic"]

    @property
    def fio(self) -> str:
        """
        Возвращает:
            ФИО одной строкой, из непустых частей.
        """
        return " ".join(filter(None, [self.surname, self.name, self.patronymic]))

    @property
    def method_label(self) -> str:
        """
        Возвращает:
            Человекочитаемая подпись запланированного способа голосования.
        """
        return METHOD_LABELS.get(self.method, "")

    @property
    def voted_method_label(self) -> str:
        """
        Возвращает:
            Человекочитаемая подпись фактического способа голосования.
        """
        return METHOD_LABELS.get(self.voted_method, "")

    def __str__(self) -> str:
        """
        Возвращает:
            Строковое представление для админки и списков: "Таб№ ФИО".
        """
        return f"{self.tab_number} {self.fio}"
