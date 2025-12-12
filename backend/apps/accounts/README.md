# Accounts App

## Назначение

Приложение `accounts` отвечает за управление пользователями и ролями в системе KindEltern Web.

## Основные задачи

- **Регистрация и управление пользователями** — создание, редактирование, деактивация учётных записей.
- **Аутентификация** — логин/логаут, управление сессиями.
- **Ролевое управление доступом**:
  - `ADMIN` — ведение картотеки, синхронизация, работа с отклонёнными изменениями.
  - `OPERATOR` — ограниченный Admin (SEPA, прошлые месяцы).
  - `SUPERADMIN` — утверждение/отклонение рискованных изменений.
  - `USER` — только просмотр записей и истории (read-only).
- **User Cabinet** — отдельный интерфейс для роли User.

## Ключевые сущности

- `User` — расширение стандартной Django-модели с полем роли (`UserRole` enum).
- `UserRole` — enum с ролями: `ADMIN`, `OPERATOR`, `SUPERADMIN`, `USER`.

## User Cabinet (PROMPT_10)

Веб-интерфейс для роли User (read-only доступ):

### Views (`views.py`)

- `role_based_redirect()` — функция редиректа на корневом URL `/` в зависимости от роли.
- `UserDashboardView` — стартовая страница (`/user/`), форма поиска, информация о правах.
- `UserKarteiSearchView` — поиск записей (`/user/search/`), фильтры: FamilyID, Parent, Child, Year.
- `UserKarteiDetailView` — детали записи (`/user/record/<pk>/`), read-only.
- `UserRecordHistoryView` — история записи (`/user/record/<pk>/history/`), аналог Geschichte.

### Permission Mixins

- `UserRoleMixin` — ограничивает доступ только для роли `USER`. Блокирует все POST/PUT/DELETE.
- `AnyAuthenticatedUserMixin` — для views, доступных всем авторизованным.

### URLs (`urls.py`)

```
/user/                      — User Dashboard
/user/search/               — Search records
/user/record/<pk>/          — Record detail (read-only)
/user/record/<pk>/history/  — Record history
```

### Templates

- `accounts/user_dashboard.html` — стартовая страница.
- `accounts/user_search.html` — поиск и результаты.
- `accounts/user_record_detail.html` — детали записи.
- `accounts/user_record_history.html` — история изменений.

## Связь с VBA-модулями

Роли соответствуют файлам:

- `Admin` → `KindElternDaten_XX_Admin.xlsm`
- `Superadmin` → `KindElternDaten_XX_Suprime.xlsm`
- `User` → нет прямого аналога (новая роль для веб-системы).
