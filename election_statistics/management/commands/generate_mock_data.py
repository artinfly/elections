# election_statistics/management/commands/generate_mock_data.py
import random

from django.core.management.base import BaseCommand

from election_statistics.models import DEG, UIK, UIK19, UVZ, Employee


class Command(BaseCommand):
    help = "Генерирует 1000 тестовых сотрудников с разными комбинациями округов и УИК"

    def handle(self, *args, **options):
        self.stdout.write("Очистка базы данных...")
        Employee.objects.all().delete()

        productions = ["Механосборочное", "Инструментальное", "Литейное", "Кузнечное"]
        services = ["ИТР", "Рабочие", "Администрация", "Служба качества"]
        surnames = [
            "Иванов",
            "Петров",
            "Сидоров",
            "Смирнов",
            "Кузнецов",
            "Попов",
            "Васильев",
            "Зайцев",
            "Соколов",
            "Михайлов",
        ]
        names = [
            "Иван",
            "Петр",
            "Сергей",
            "Андрей",
            "Алексей",
            "Дмитрий",
            "Владимир",
            "Николай",
        ]
        patronymics = [
            "Иванович",
            "Петрович",
            "Сергеевич",
            "Андреевич",
            "Алексеевич",
            "Дмитриевич",
        ]

        okrugs_options = ["19", "20", "21", ""]
        okrugs_weights = [15, 40, 40, 5]  # Пустых округов мало

        batch = []
        self.stdout.write("Генерация 1000 сотрудников...")

        for i in range(1, 1001):
            okrug = random.choices(okrugs_options, weights=okrugs_weights)[0]

            # Логика УИК: если округ 19, то УИК начинается на 19. Если 20/21 - на 20/21.
            if okrug == "19":
                uik = f"19{random.randint(10, 99)}"
            elif okrug in ["20", "21"]:
                uik = f"{okrug}{random.randint(10, 99)}"
            else:
                uik = ""

            # Специальная логика для "Не 19 округ" (Открепившиеся):
            # Люди из 20/21 округа, но прикрепленные к УИК 19 (помечаем detached=True)
            detached = False
            if okrug in ["20", "21"] and random.random() < 0.05:
                uik = f"19{random.randint(10, 99)}"
                detached = True

            if okrug == "19":
                method = random.choices(
                    [UIK19, DEG, UIK, ""], weights=[60, 20, 10, 10]
                )[0]
            else:
                method = random.choices([DEG, UIK, UVZ, ""], weights=[30, 40, 20, 10])[
                    0
                ]

            voted = random.random() < 0.65
            voted_method = method if voted and method else ""
            not_going = not voted and random.random() < 0.15

            person = Employee(
                tab_number=f"{i:06d}",
                department=f"{random.randint(1, 50):03d}",
                production=random.choice(productions) if random.random() > 0.1 else "",
                service=random.choice(services),
                surname=random.choice(surnames),
                name=random.choice(names),
                patronymic=random.choice(patronymics),
                position="Специалист",
                category=random.choice(["Рабочий", "ИТР", "Служащий"]),
                birth_date=f"1980-{random.randint(1,12):02d}-{random.randint(1,28):02d}",
                region="Свердловская",
                city="Нижний Тагил",
                street="Ленина",
                house=str(random.randint(1, 100)),
                uik=uik,
                uik_address="Школа",
                district="Дзержинский",
                okrug=okrug,
                method=method,
                voted=voted,
                voted_method=voted_method,
                detached=detached,
                not_going=not_going,
            )
            batch.append(person)

            if len(batch) >= 200:
                Employee.objects.bulk_create(batch)
                batch = []

        if batch:
            Employee.objects.bulk_create(batch)

        self.stdout.write(
            self.style.SUCCESS("Успешно сгенерировано 1000 тестовых сотрудников!")
        )
