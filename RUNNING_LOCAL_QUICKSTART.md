# Быстрый старт KindEltern Web (локальный Docker)

Этот файл дополняет `RUNNING.md` и даёт пошаговый сценарий для локального развёртывания и ручного тестирования ролей/функционала.

## 1. Предварительные требования

- Установлены:
  - Docker (Docker Desktop / Engine)
  - Docker Compose v2 (`docker compose version`)
- Порт `8000` свободен на локальной машине.

## 2. Подготовка окружения (`.env`)

В корне репозитория:

```bash
cp .env.example .env
```

Далее отредактируйте `.env`:

- `DJANGO_SECRET_KEY` — задайте произвольный длинный ключ (можно сгенерировать командой из комментария в `.env.example`).
- `DJANGO_DEBUG=True` — оставить для разработки.
- `POSTGRES_PASSWORD` — задайте непустой пароль (например, `kindeltern_local`).

Переменные `ACCESS_BASE_DIR` / `ACCESS_CONN_STRING_TEMPLATE` можно пока не трогать, если импорт из Access не нужен.

## 2.1 Импорт из Access на Windows (важный нюанс про шаблон ODBC)

Импорт `*.accdb` почти всегда удобнее запускать **на Windows‑хосте** (не внутри Docker), потому что драйвер `Microsoft Access Driver (*.mdb, *.accdb)` относится к Windows (ACE / Access Database Engine).

В `.env`:

- `ACCESS_BASE_DIR` должен указывать на каталог с файлами `.accdb` (например, `c:\Konst\...\Base-legasy`).
- `ACCESS_CONN_STRING_TEMPLATE` должен содержать плейсхолдер `{file_path}` и **удвоенные фигурные скобки** вокруг имени драйвера:

```env
ACCESS_CONN_STRING_TEMPLATE=DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};
```

Это важно, потому что код подставляет `{file_path}` через Python `.format(...)`. Если оставить одинарные `{...}` вокруг DRIVER, будет `KeyError: 'Microsoft Access Driver (*.mdb, *.accdb)'` и подключение к Access не произойдёт.

## 3. Сборка и запуск контейнеров

Из корня проекта:

```bash
docker compose build web
docker compose up -d
```

Проверка состояния:

```bash
docker compose ps
docker compose logs web --tail=50
```

Ожидаемые строки в логах `web`:

- `Waiting for database...`
- `Database is ready!`
- `Applying database migrations...`
- `Starting Django development server...`

После этого приложение доступно по адресу `http://localhost:8000/`.

## 3.1 Автоматические проверки разработчика

Из корня репозитория:

```bash
python backend/manage.py check
python backend/manage.py makemigrations --check --dry-run
pytest -m "not browser" --reuse-db
```

Для automated browser smoke tests после `pip install -r backend/requirements.txt` один раз установите Chromium binary:

```bash
python -m playwright install chromium
```

Запуск минимального browser smoke baseline:

```bash
pytest -m browser backend/tests/browser --browser chromium --reuse-db
```

Эти тесты используют `pytest-playwright` + `pytest-django live_server` и создают fixture-backed users/records в тестовой БД. Они покрывают login/root redirects по ролям, ADMIN Kartei create/edit entry points, OPERATOR redirects для `PENDING`/`DECLINED`, SUPERADMIN pending overview + War/Ist detail и USER read-only cabinet/detail route. Destructive approve/decline/bulk apply и catalog group pricing smoke здесь намеренно не автоматизированы.

Не запускайте backend pytest и browser pytest параллельно против одной PostgreSQL test DB. Если после оборванного запуска появляется `database "test_kindeltern" already exists`, остановите лишние pytest/live_server процессы и очистите stale test DB:

```bash
docker compose exec db sh -lc 'TEST_DB="test_${POSTGRES_DB:-kindeltern}"; psql -U "$POSTGRES_USER" -d postgres -c "SELECT pg_terminate_backend(pid) FROM pg_stat_activity WHERE datname = '"'"'${TEST_DB}'"'"';"; dropdb -U "$POSTGRES_USER" --if-exists "$TEST_DB"'
```

После migration changes используйте `--create-db`; для повторных локальных прогонов без migration changes используйте `--reuse-db`.

> Если в логах постоянно повторяется только `Database not ready, waiting...`, убедитесь, что:
>
> - изменённый `backend/entrypoint.sh` действительно попал в образ (повторный `docker compose build web`);
> - контейнер `db` имеет статус `healthy` (`docker compose ps`).

## 4. Создание суперпользователя Django

Суперпользователь нужен для входа в Django Admin и управления пользователями:

```bash
docker compose exec web python manage.py createsuperuser
```

После создания:

- Откройте `http://localhost:8000/admin/`.
- Войдите под суперпользователем.

## 5. Роли и создание пользователей

Модель пользователя: `accounts.User` с полем `role`:

- `ADMIN` — полный доступ к Kartei (создание/редактирование, запуск SAFE‑изменений).
- `OPERATOR` — оператор с ограничениями по SEPA и прошлым месяцам.
- `SUPERADMIN` — утверждение/отклонение pending, NeuList, история.
- `USER` — конечный пользователь, только чтение (поиск, просмотр записей и истории).

Через Django Admin (`/admin/` → `Users`):

1. Нажмите «Add User», задайте `username` и пароль.
2. На странице редактирования пользователя установите:
   - ADMIN:
     - `role = ADMIN`
     - `is_staff = True`, `is_superuser = False`
   - OPERATOR:
     - `role = OPERATOR`
     - `is_staff = True`, `is_superuser = False`
   - SUPERADMIN:
     - `role = SUPERADMIN`
     - `is_staff = True`, `is_superuser = False` (или True, если хотите объединить с правами superuser’а)
   - USER:
     - `role = USER`
     - `is_staff = False`, `is_superuser = False`
3. Убедитесь, что `is_active = True` для всех.

## 6. Ручной сценарий тестирования

### 6.1 SUPERADMIN

- Войти под `role=SUPERADMIN` на `http://localhost:8000/`:
  - Должно произойти перенаправление на `/approvals/superadmin/pending/`.
  - Проверить:
    - `/approvals/superadmin/pending/` — обзор pending‑изменений.
    - `/approvals/superadmin/neu/` — NeuList (новые записи).
    - `/approvals/superadmin/history/<id>/` — история по записи.

### 6.2 ADMIN / OPERATOR

- Войти под `role=ADMIN` или `role=OPERATOR`:
  - Корневой `/` → `/karteien/` (список Kartei).
  - Через UI создать несколько `KarteiRecord` для текущего года.
  - Проверить:
    - Изменение не‑tracked полей (SAFE) приводит к прямому обновлению записи.
    - Изменение tracked‑полей (FamilyID, Parent, Child, месяцы, Extra и т.д.) создаёт `PendingChange`, меняет статус записи на `PENDING`.
  - Для OPERATOR:
    - Проверить, что записи с SEPA и прошлые месяцы блокируются согласно правилам (предупреждения/запреты).

### 6.3 Цепочка APPROVE / DECLINE

1. Под ADMIN/OPERATOR внести рискованное изменение (tracked‑поле) → запись `PENDING`.
2. Под SUPERADMIN:
   - Открыть `/approvals/superadmin/pending/`, убедиться, что pending появился.
   - Перейти в War/Ist (`/approvals/superadmin/pending/<pending_id>/`).
3. Принять решение:
   - **Approve**:
     - запись становится `NORMAL`;
     - изменения применены к KarteiRecord;
     - в истории появляется событие `APPROVE`;
     - Admin’ам приходит уведомление `APPROVED`.
   - **Decline**:
     - запись становится `DECLINED`, создаётся `DeclinedChange` с причиной;
     - запись видна в `/approvals/declined/`;
     - в истории — `DECLINE`;
     - Admin’ам приходит уведомление `DECLINED_CREATED`.

### 6.4 USER (read‑only)

- Войти под `role=USER`:
  - Корневой `/` → `/user/` (User Dashboard).
  - Проверить:
    - `/user/search/` — поиск записей по FamilyID/Parent/Child/Year.
    - `/user/record/<pk>/` — просмотр записи без кнопок редактирования.
    - `/user/record/<pk>/history/` — просмотр истории по записи.
  - Убедиться, что никакие действия пользователя не изменяют данные (только GET‑запросы, без форм сохранения).

## 7. Остановка и сброс

- Остановить сервисы, сохранив данные:

```bash
docker compose down
```

- Полный сброс (с удалением данных Postgres):

```bash
docker compose down -v
```

## 8. Структура ключей KarteiRecord и импорт нескольких лет

Django PK (`pkid`) — суррогатный ключ `BigAutoField`, используется для FK-связей.  
Доменный ключ — `(year, id)`:

- `id` — Access ID (поле AV / tblKartei.ID) — **не уникален глобально!**
- `year` — год записи (2024, 2025, ...)

Это позволяет хранить данные нескольких лет одновременно.

### Переимпорт данных после обновления модели

Если вы обновили модель KarteiRecord (например, после применения PROMPT_12), необходимо:

1. **Сбросить базу данных:**

   ```powershell
   docker compose down -v
   ```

2. **Пересобрать и запустить:**

   ```powershell
   docker compose up -d --build
   ```

3. **Импортировать годы последовательно (из Windows, не из Docker!):**

   > **Рекомендация:** Импорт Access лучше запускать из локальной командной строки Windows,  
   > поскольку ODBC-драйвер для Access — это Windows-компонент.

   ```powershell
   # Подготовка окружения (один раз)
   cd backend
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt

   # Установить переменные окружения
   $env:DATABASE_URL = "postgres://kindeltern:kindeltern_local@localhost:5432/kindeltern"
   $env:ACCESS_BASE_DIR = "C:\Path\To\AccessFiles"
   $env:ACCESS_CONN_STRING_TEMPLATE = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};"

   # Импорт 2024
   python manage.py import_access_year `
       --year 2024 `
       --access-file KindElternDaten_24_front.accdb `
       --report-dir ..\import_reports

   # Импорт 2025
   python manage.py import_access_year `
       --year 2025 `
       --access-file KindElternDaten_25_front.accdb `
       --report-dir ..\import_reports
   ```

4. (Опционально) **Patch-доимпорт отдельных полей** (без полного реимпорта).

   Используйте, если нужно дописать в уже импортированные записи поля, которые раньше не импортировались
   (учителя Value11/Value16, маркеры контракта Value14/Value20, `sepa_marker` Value47).

   > **Важно (миграции):** перед `--patch-fields` убедитесь, что миграции применены к этой базе данных,  
   > иначе будет ошибка вида `UndefinedColumn ... teacher1_legacy_name does not exist`.  
   > - если Django запущен в Docker: `docker compose exec web python manage.py migrate`  
   > - если импорт запускаете локальным Python: `python manage.py migrate` (при том же `DATABASE_URL`)

   ```powershell
   # Patch-доимпорт 2025 (пример)
   python manage.py import_access_year `
       --year 2025 `
       --access-file KindElternDaten_25_front.accdb `
       --patch-fields `
       --report-dir ..\import_reports
   ```

   > **Важно:** Убедитесь, что контейнер `db` запущен (`docker compose up -d db`),  
   > чтобы PostgreSQL был доступен на `localhost:5432`.

После импорта в БД могут существовать записи с одинаковым Access `id`, но разными `year`.
