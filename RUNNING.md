# Running KindEltern Web

Этот документ описывает, как запустить проект локально для разработки и как развернуть его на VPS.

---

## Требования

- **Docker** (версия 20.10+)
- **Docker Compose** (v2.0+ или `docker compose` plugin)

---

## Локальный запуск (Development)

### 1. Подготовка окружения

1. Скопируйте файл с переменными окружения:

   ```bash
   cp .env.example .env
   ```

2. Отредактируйте `.env` при необходимости:
   - Для локальной разработки значения по умолчанию работают "из коробки".
   - При желании измените `POSTGRES_PASSWORD` на более безопасный.

### 2. Запуск проекта

```bash
docker compose up --build
```

Эта команда:

- Соберёт Docker-образ для Django-приложения.
- Запустит PostgreSQL контейнер.
- Дождётся готовности базы данных.
- Применит миграции Django.
- Запустит development-сервер.

### 3. Проверка работы

- Приложение доступно по адресу: **http://localhost:8000/**
- Django Admin: **http://localhost:8000/admin/**
- База данных PostgreSQL доступна на порту `5432` (для отладки с pgAdmin, DBeaver и т.п.)

### 4. Остановка

```bash
# Остановить контейнеры (данные сохраняются)
docker compose down

# Остановить и удалить данные (volume с БД)
docker compose down -v
```

### 5. Полезные команды

```bash
# Пересобрать образ после изменений в requirements.txt или Dockerfile
docker compose up --build

# Запустить в фоновом режиме
docker compose up -d

# Посмотреть логи
docker compose logs -f web
docker compose logs -f db

# Выполнить команду внутри контейнера
docker compose exec web python manage.py createsuperuser
docker compose exec web python manage.py shell
docker compose exec web python manage.py makemigrations
```

---

## Деплой на VPS (Production)

### 1. Подготовка сервера

1. Установите Docker и Docker Compose:

   ```bash
   # Ubuntu/Debian
   sudo apt update
   sudo apt install -y docker.io docker-compose-plugin
   sudo systemctl enable docker
   sudo systemctl start docker

   # Добавить пользователя в группу docker (опционально)
   sudo usermod -aG docker $USER
   ```

2. Клонируйте репозиторий на сервер:

   ```bash
   git clone <your-repo-url> /opt/kindeltern
   cd /opt/kindeltern
   ```

### 2. Настройка окружения

1. Создайте файл `.env` с production-настройками:

   ```bash
   cp .env.example .env
   nano .env  # или vim .env
   ```

2. **Обязательные изменения для production:**

   ```env
   # Сгенерируйте новый секретный ключ!
   DJANGO_SECRET_KEY=<сгенерированный-ключ>

   # Отключите debug-режим
   DJANGO_DEBUG=False

   # Укажите ваш домен
   DJANGO_ALLOWED_HOSTS=example.com,www.example.com

   # Для HTTPS (если используете)
   DJANGO_CSRF_TRUSTED_ORIGINS=https://example.com,https://www.example.com

   # Установите надёжный пароль для БД
   POSTGRES_PASSWORD=<надёжный-пароль>
   ```

3. Сгенерируйте секретный ключ:

   ```bash
   docker compose run --rm web python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"
   ```

### 3. Запуск на сервере

```bash
# Запуск в фоновом режиме
docker compose up -d --build
```

Приложение будет доступно на порту 8000.

### 4. Создание суперпользователя

```bash
docker compose exec web python manage.py createsuperuser
```

### 5. Мониторинг и логи

```bash
# Статус контейнеров
docker compose ps

# Логи приложения
docker compose logs -f web

# Логи базы данных
docker compose logs -f db
```

### 6. Обновление приложения

```bash
cd /opt/kindeltern
git pull origin main
docker compose up -d --build
```

---

## Продакшн-рекомендации

На данном этапе настроен базовый запуск. Для полноценного production-окружения рекомендуется:

1. **Reverse Proxy (nginx/traefik):**

   - Проксирование запросов на порт 8000.
   - Обработка SSL/TLS сертификатов.
   - Статические файлы.

2. **HTTPS:**
   - Let's Encrypt через certbot или traefik.
3. **Gunicorn вместо runserver:**

   - Уже добавлен в `requirements.txt`.
   - Для запуска с Gunicorn измените команду:

   ```yaml
   # В docker-compose.yml для web-сервиса добавьте:
   command: gunicorn
   ```

4. **Бэкапы базы данных:**

   - Настроить регулярное резервное копирование volume `postgres_data`.

5. **Мониторинг:**
   - Prometheus, Grafana, или простой healthcheck endpoint.

---

## Структура файлов Docker

```
.
├── .env.example           # Пример переменных окружения
├── .env                   # Реальные значения (не в git!)
├── docker-compose.yml     # Оркестрация контейнеров
└── backend/
    ├── Dockerfile         # Сборка Django-образа
    ├── entrypoint.sh      # Скрипт запуска (миграции + сервер)
    └── requirements.txt   # Python-зависимости
```

---

## Troubleshooting

### Ошибка "Database not ready"

База данных ещё не успела запуститься. `entrypoint.sh` автоматически ждёт готовности БД, но если ожидание слишком долгое:

```bash
# Проверьте статус контейнера db
docker compose logs db
```

### Ошибка "Permission denied" для entrypoint.sh

На Windows файл может потерять executable-флаг. Пересоберите образ:

```bash
docker compose build --no-cache web
```

### Миграции не применяются

Миграции применяются автоматически при каждом запуске через `entrypoint.sh`. Для ручного применения:

```bash
docker compose exec web python manage.py migrate
```

---

## Импорт данных из Access (legacy_import)

Этот раздел описывает, как импортировать данные из Access-базы (`*.accdb`) в PostgreSQL.

### Предварительные требования

1. **pyodbc установлен** в Python-окружении backend (уже в `requirements.txt`).
2. **ODBC-драйвер для Access:**
   - **Windows:** Microsoft Access Database Engine / ACE ODBC Driver.
   - **Linux:** связка `unixODBC` + MDB Tools или аналог (см. `backend/apps/legacy_import/README.md`).

### Где лежат `.accdb`

Укажите расположение файлов через переменные окружения или `settings.py`:

```bash
# Базовый каталог с .accdb файлами
ACCESS_BASE_DIR=/opt/kindeltern_data

# Шаблон строки подключения ODBC
ACCESS_CONN_STRING_TEMPLATE="DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ={file_path};"
```

Файлы должны называться по шаблону `KindElternDaten_XX_front.accdb` (где XX — год).

### Dry-run импорт (анализ без записи)

```bash
docker compose exec web python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --dry-run \
    --report-dir ./import_reports
```

В режиме `--dry-run`:

- База данных **не изменяется**.
- Генерируются JSON/CSV-отчёты (статистика, FamilyID-issues, возможный merge).

### Боевой импорт

```bash
docker compose exec web python manage.py import_access_year \
    --year 2025 \
    --access-file KindElternDaten_25_front.accdb \
    --familyid-policy=report \
    --sync-history \
    --report-dir ./import_reports
```

Основные параметры:

| Параметр            | Описание                                                     |
| ------------------- | ------------------------------------------------------------ |
| `--year`            | Год для импортируемых записей (обязательный)                 |
| `--access-file`     | Имя файла или путь к `.accdb` (обязательный)                 |
| `--dry-run`         | Только анализ, без записи в БД                               |
| `--familyid-policy` | `report` (по умолчанию) или `auto-merge` — политика FamilyID |
| `--sync-history`    | Синхронизировать `history_raw` → `HistoryEvent`              |
| `--skip-pending`    | Пропустить импорт `pre_tblKartei` (pending-изменения)        |
| `--skip-declined`   | Пропустить импорт `decl_tblKartei` (declined-изменения)      |
| `--report-dir`      | Каталог для CSV/JSON-отчётов                                 |

### Где смотреть результат

После импорта в `--report-dir` создаются:

- `import_stats_YEAR.json` — полная статистика импорта.
- `import_familyid_issues_YEAR.csv` — обнаруженные проблемы FamilyID.
- `import_familyid_merge_YEAR.csv` — маппинг авто-нормализации (если был `auto-merge`).
