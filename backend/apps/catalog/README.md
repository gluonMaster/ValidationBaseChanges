# Catalog App (Reference Data)

Приложение `catalog` содержит справочные таблицы для стандартизации данных о преподавателях, предметах и ценах.

## Назначение

Эти таблицы **не заменяют** существующие поля `KarteiRecord.subject1/subject2` и не влияют на импорт из Access. Это отдельные справочники для:

- Стандартизации названий предметов и имён преподавателей
- Связи "кто вёл что в каком году"
- Управления прайс-листом по годам и дисциплинам
- Будущих форм выбора с автодополнением
- Отчётности и аналитики

## Модели

### Teacher (Преподаватель)

| Поле       | Тип          | Описание               |
| ---------- | ------------ | ---------------------- |
| last_name  | CharField    | Фамилия (обязательное) |
| first_name | CharField    | Имя (обязательное)     |
| is_active  | BooleanField | Активен (default=True) |

**Ограничение:** уникальность по `(last_name, first_name)`.

### Subject (Предмет)

| Поле      | Тип          | Описание                       |
| --------- | ------------ | ------------------------------ |
| name      | CharField    | Название предмета (уникальное) |
| is_active | BooleanField | Активен (default=True)         |

### TeachingAssignment (Назначение на курс)

| Поле      | Тип                       | Описание                      |
| --------- | ------------------------- | ----------------------------- |
| year      | PositiveSmallIntegerField | Учебный год (2024, 2025, ...) |
| subject   | FK → Subject              | Предмет                       |
| teacher   | FK → Teacher              | Преподаватель                 |
| is_active | BooleanField              | Активно (default=True)        |

**Ограничение:** уникальность по `(year, subject, teacher)`.

**Индексы:**

- `(year, subject)` — для быстрого поиска преподавателей по году и предмету
- `(year, teacher)` — для быстрого поиска предметов по году и преподавателю

### PriceOption (Прайс-лист)

| Поле      | Тип                       | Описание                                          |
| --------- | ------------------------- | ------------------------------------------------- |
| year      | PositiveSmallIntegerField | Учебный год (2024, 2025, ...)                     |
| subject   | FK → Subject              | Предмет                                           |
| amount    | DecimalField(10,2)        | Сумма в € (≥ 0)                                   |
| comment   | TextField                 | Комментарий — почему такая цена (blank, optional) |
| is_active | BooleanField              | Активна ли цена (default=True)                    |

**Ограничение:** уникальность по `(year, subject, amount, comment)`.

**Индексы:**

- `(year, subject)` — для быстрого поиска цен по году и предмету

#### Семантика цены (`amount`)

Поле `amount` имеет разную семантику в зависимости от типа занятий:

| Тип занятий                                           | Единица | Пример предметов                 |
| ----------------------------------------------------- | ------- | -------------------------------- |
| Групповые (по умолчанию)                              | €/Monat | Gitarre, Klavier, Musiktheorie   |
| Индивидуальные (в названии "Ind." или "VSpE\_")       | €/UE    | Ind. Gitarre, VSpE_Klavier       |
| Наххильфе/репетиторство (в названии "NH"/"Nachhilfe") | €/UE    | NH Mathematik, Nachhilfe Deutsch |

**UE** = Unterrichtseinheit (академический час, 45 минут).

Методы модели:

- `get_price_unit()` → `"€/Monat"` или `"€/UE"` (автоматически определяется по названию предмета)
- `is_per_hour()` → `True` если цена за академический час

## Скидки (Discounts)

### Discount (Справочник скидок)

| Поле        | Тип          | Описание                                     |
| ----------- | ------------ | -------------------------------------------- |
| kind        | CharField    | Тип: `PERCENT` (%) или `FIXED` (€)           |
| value       | DecimalField | Для PERCENT: 0.00-0.99. Для FIXED: сумма в € |
| description | TextField    | Описание скидки (blank)                      |
| is_active   | BooleanField | Активна ли скидка (default=True)             |

**Примеры:**

- `kind=PERCENT, value=0.25` → скидка 25%
- `kind=FIXED, value=10.00` → скидка 10€

### FamilyDiscount (Скидка на семью)

| Поле        | Тип                       | Описание                            |
| ----------- | ------------------------- | ----------------------------------- |
| year        | PositiveSmallIntegerField | Учебный год                         |
| family_id   | CharField                 | Идентификатор семьи                 |
| discount    | FK → Discount             | Применяемая скидка                  |
| start_month | PositiveSmallIntegerField | Начальный месяц (1-12), default=1   |
| end_month   | PositiveSmallIntegerField | Конечный месяц (1-12), default=12   |
| months      | JSONField                 | Опционально: список месяцев [1,3,5] |
| created_at  | DateTimeField             | Время создания                      |
| updated_at  | DateTimeField             | Время обновления                    |

Применяется ко всем записям `KarteiRecord` с данным `family_id` в указанном году.

### RecordDiscount (Скидка на запись)

| Поле        | Тип                       | Описание                             |
| ----------- | ------------------------- | ------------------------------------ |
| record      | FK → KarteiRecord         | Запись, к которой применяется скидка |
| discount    | FK → Discount             | Применяемая скидка                   |
| start_month | PositiveSmallIntegerField | Начальный месяц (1-12), default=1    |
| end_month   | PositiveSmallIntegerField | Конечный месяц (1-12), default=12    |
| months      | JSONField                 | Опционально: список месяцев [1,3,5]  |
| created_at  | DateTimeField             | Время создания                       |
| updated_at  | DateTimeField             | Время обновления                     |

Применяется к конкретной записи `KarteiRecord`.

### Логика месяцев

Метод `get_applicable_months()` возвращает список месяцев:

- Если `months` заполнено → используется этот список
- Иначе → диапазон `start_month..end_month`

```python
# Примеры
fd = FamilyDiscount(start_month=1, end_month=12)  # весь год
fd.get_applicable_months()  # → [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]

rd = RecordDiscount(months=[1, 6, 12])  # только конкретные месяцы
rd.get_applicable_months()  # → [1, 6, 12]
```

## Примеры запросов

```python
from apps.catalog.models import Teacher, Subject, TeachingAssignment, PriceOption

# Все преподаватели, которые вели "Gitarre" в 2024 году
teachers = Teacher.objects.filter(
    assignments__year=2024,
    assignments__subject__name="Gitarre",
    assignments__is_active=True
)

# Все предметы преподавателя "Müller, Hans" за все годы
subjects = Subject.objects.filter(
    assignments__teacher__last_name="Müller",
    assignments__teacher__first_name="Hans"
).distinct()

# Список всех назначений за 2025 год
assignments_2025 = TeachingAssignment.objects.filter(
    year=2025,
    is_active=True
).select_related('subject', 'teacher')

# Все годы, в которые велся предмет "Klavier"
years = TeachingAssignment.objects.filter(
    subject__name="Klavier"
).values_list('year', flat=True).distinct()

# Все активные цены для предмета в 2025 году
prices = PriceOption.objects.filter(
    year=2025,
    subject__name="Gitarre",
    is_active=True
).order_by('amount')

# Вывод цен с единицами
for price in prices:
    unit = price.get_price_unit()
    print(f"{price.amount} {unit}: {price.comment or 'Standardpreis'}")
```

## Django Admin

Все четыре модели зарегистрированы в Django Admin (`/admin/`):

- **Teachers**: поиск по фамилии/имени, фильтр по активности
- **Subjects**: поиск по названию, фильтр по активности
- **Teaching Assignments**: фильтры по году/предмету/преподавателю, autocomplete для выбора
- **Price Options**: фильтры по году/предмету/активности, отображение единицы цены

## Web UI

Админ-интерфейс доступен по адресу `/catalog/` (требуется роль Admin):

- `/catalog/` — главная страница со статистикой
- `/catalog/teachers/` — список и управление преподавателями
- `/catalog/subjects/` — список и управление предметами
- `/catalog/assignments/` — назначения преподаватель-предмет-год
- `/catalog/assignments/copy-year/` — копирование назначений между годами
- `/catalog/prices/` — прайс-лист по годам и предметам
- `/catalog/prices/copy-year/` — копирование прайса между годами
- `/catalog/discounts/` — справочник скидок (процентные/фиксированные)
- `/catalog/family-discounts/` — скидки на семьи (по году и family_id)
- `/catalog/record-discounts/` — скидки на конкретные записи (по pkid)

### Копирование года

Функция "Jahr kopieren" позволяет:

- Копировать все назначения или цены из одного года в другой
- Опция "Nur aktive" — копировать только активные записи
- Опция "Überschreiben" — деактивировать существующие записи в целевом году

## Миграции

```bash
cd backend
python manage.py makemigrations catalog
python manage.py migrate
```
