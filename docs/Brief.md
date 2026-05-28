# Brief - KindEltern Web

## Что строим

KindEltern Web - веб-система (Django + PostgreSQL), заменяющая связку Excel+VBA+Access (KindElternDaten*XX*\*.xlsm и KindElternDaten_XX_front.accdb).

## Зачем

- Единый источник истины: централизованная БД вместо набора файлов по годам.
- Удобный UI для ежедневной работы (без Excel/VBA).
- Контролируемые правки через approvals workflow (pending/declined).
- Понятная история изменений и аудит.
- Автоматизация начислений по месяцам с учетом предметов, прайса и скидок.
- Возможность импортировать и доимпортировать (patch) данные из Access.

## Для кого (роли)

- `ADMIN`: ведение записей (семьи/дети), назначение скидок, работа с declined, мастер новой семьи.
- `OPERATOR`: почти как ADMIN, но с ограничениями (например, без delete).
- `SUPERADMIN`: принятие решений по pending (approve/decline) + read-only просмотр `/karteien/`.
- `USER`: read-only просмотр записей/истории и read-only `/karteien/` (list/detail/live search).

## Основные артефакты проекта

- Архитектура: `ARCHITECTURE.md`
- Модель данных/маппинг Access: `DOMAIN_MODEL.md`
- Короткая карта проекта (для экономии контекста): `docs/ProjectMap.md`
- Пайплайн разработки через промпты: `PROMPTS_OVERVIEW.md`
- Handoff для нового чата: `PROMPT_00_NEXT_CHAT_HANDOFF.md`
- Текущее состояние/подводные камни: `docs/Status.md`, `docs/Spec.md`, `docs/Constraints.md`

## Текущее состояние (супер-коротко)

- Применены изменения до `PROMPT_109` включительно.
- NeuList для Superadmin работает per-year (доменный ключ `(year, id)`), а Historie отображается как таймлайн (не только raw-строкой).
