# Подготовка к работе с Excel и развёртывание на сервере

## 1. Хранение данных на сервере

### Текущее состояние
- Данные лежат в JSON-файлах в `public/`: `schedule.json`, `attendance.json`, `students.json`, `vedomost.json`, `schedule_grid.json`.
- Бэкенд читает их через `data.LoadJSON`; кэш по времени модификации файла.
- Подходит для демо и небольшой нагрузки.

### Рекомендации для продакшена

| Вариант | Плюсы | Минусы |
|--------|--------|--------|
| **Оставить JSON** | Уже работает, не нужна БД | Нет транзаций, сложнее параллельная запись, нет истории |
| **SQLite** | Один файл, не нужен отдельный сервер БД | Ограничения при высокой конкурентной записи |
| **PostgreSQL** | Надёжность, бэкапы, масштаб | Нужна установка и администрирование БД |

**Предлагаемая схема БД (если переходить с JSON):**

- `departments` (id, name)
- `groups` (id, department_id, name)
- `students` (id, group_id, full_name, number_in_group, status)
- `schedule_records` (id, group_id, date, lesson_number, discipline, teacher, attendance)
- `attendance_records` (id, department, group, student, date, missed) — или ссылки на student_id
- `teachers`, `disciplines` — при необходимости нормализации

Связи: по `group_id`, `student_id`; даты в едином формате (например `DATE`).

---

## 2. Загрузка Excel и обновление данных

### Текущая архитектура
- **Расписание:** Excel (`расписание.xls`) → конвертер → `schedule_grid.json` и/или `schedule.json` (через `ConvertScheduleGridToLessonsFormat` + `students.json`).
- **Ведомость:** Excel (`vedomost.xls`) → конвертер → `vedomost.json`.
- **Мастер-конвертация:** один Excel (ведомость) → `students.json`, `attendance.json`, `vedomost.json` (см. `cmd/convert-master`).

Эндпоинты (админ):
- `POST /api/admin/convert/statement` — загрузка файла ведомости → обновление `vedomost.json`.
- `POST /api/admin/convert/schedule` — загрузка расписания → обновление `schedule.json`.
- `POST /api/admin/refresh-data` — полное обновление данных (конвертеры по конфигу).

### Что нужно для автоматизации
1. **Единая точка входа** — один endpoint типа `POST /api/admin/upload` с типом файла (расписание / ведомость / контингент) или отдельные endpoint’ы под каждый тип.
2. **Очередь задач** — при больших файлах конвертировать в фоне и отдавать статус (уже есть заготовка refresh-status).
3. **Валидация** — проверка структуры листов и обязательных колонок до записи в JSON/БД.
4. **Бэкапы** — перед перезаписью JSON копировать текущие файлы в `backup/` с датой в имени.

### Развёртывание на сервере
1. Собрать бэкенд: `go build -o server ./cmd/server`.
2. Собрать фронт: `npm run build`, положить `dist/` в статику бэкенда или отдавать через nginx.
3. Вынести пути к JSON и к загружаемым файлам в переменные окружения (уже частично через config).
4. Запуск: `./server` или через systemd/supervisor; при использовании БД — миграции перед первым запуском.
5. Nginx: проксирование `/api` на бэкенд; раздача статики фронта с корня.
6. HTTPS и JWT — секреты в env, не в репозитории.

---

## 3. Конвертация Excel → JSON (как сейчас)

| Вход | Выход | Команда / код |
|------|--------|----------------|
| ведомость.xls | students.json, attendance.json, vedomost.json | `convert-master` или Python-скрипт из конфига |
| расписание.xls | schedule_grid.json, schedule.json | `convert-schedule`, конвертер сетки + объединение с students |
| vedomost.xls | vedomost.json | `convert-statement` |

Чтобы система обновлялась по загрузке Excel:
- Колледж выгружает файлы в согласованном формате (см. список запросов к колледжу).
- Админ загружает файлы через админ-панель (или по API).
- Бэкенд запускает нужный конвертер и перезаписывает JSON (или таблицы БД при переходе на БД).
