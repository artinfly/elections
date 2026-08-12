from django.db import models

DEG = "deg"
UIK = "uik"
UVZ = "uvz"
METHODS = [(DEG, "ДЭГ"), (UIK, "УИК"), (UVZ, "УИК-УВЗ")]
METHOD_LABELS = dict(METHODS)


class Employee(models.Model):
    tab_number = models.CharField("Таб№", max_length=20, unique=True)
    department = models.CharField("Подразделение", max_length=50, db_index=True)
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
    uik = models.CharField("УИК", max_length=10, blank=True, db_index=True)
    uik_address = models.CharField("Адрес УИК", max_length=255, blank=True)
    district = models.CharField("Район", max_length=100, blank=True)
    okrug = models.CharField("Округ", max_length=10, blank=True)
    method = models.CharField(
        "Способ голосования", max_length=3, choices=METHODS, blank=True, db_index=True
    )
    voted = models.BooleanField("Проголосовал", default=False, db_index=True)
    voted_at = models.DateTimeField("Время отметки", null=True, blank=True)

    class Meta:
        verbose_name = "Сотрудник"
        verbose_name_plural = "Сотрудники"
        ordering = ["surname", "name", "patronymic"]

    @property
    def fio(self):
        return " ".join(filter(None, [self.surname, self.name, self.patronymic]))

    @property
    def method_label(self):
        return METHOD_LABELS.get(self.method, "")

    def __str__(self):
        return f"{self.tab_number} {self.fio}"
