# Схема данных проекта «Мониторинг посещаемости»

## 1. Обзор JSON-файлов

| Файл | Назначение | Ключевые связи |
|------|------------|----------------|
| `students.json` | Контингент: отделения, группы, студенты | group, department, fullName |
| `schedule.json` | Расписание: кто на какой паре по дням | group, department, studentName, date, lessonNumber |
| `schedule_grid.json` | Сетка расписания (день/пара → группа, дисциплина, преподаватель, аудитория) | group, dayOfWeek, lessonNumber |
| `attendance.json` | Посещаемость: пропуски по датам | department, group, student, date |
| `vedomost.json` | Сводная ведомость пропусков по дисциплинам/специальностям | department, specialty, group, student (ФИО или дисциплина) |

---

## 2. Статичные поля (справочники, не меняются ежедневно)

- **Отделения** — `department` (название).
- **Группы** — `group` (код группы), привязка к отделению.
- **Студенты** — ФИО (`fullName` / `studentName` / `student`), номер в группе (`numberInGroup`, `serialNumber`), статус (`status`).
- **Преподаватели** — имена в расписании (`teacher`).
- **Дисциплины** — названия (`discipline`).
- **Аудитории** — только в `schedule_grid.json` (`location`).
- **Специальности** — только в `vedomost.json` (`specialty`).

---

## 3. Динамические поля (меняются по дате/дню)

| Источник | Поле | Описание |
|----------|------|----------|
| `schedule.json` | `date` | Дата занятия (формат `DD.MM.YYYY 0:00:00`) |
| `schedule.json` | `lessonNumber` | Номер пары (1–6) |
| `schedule.json` | `attendance` | Факт присутствия на паре (bool) |
| `attendance.json` | `date` | Дата (YYYY-MM-DD) |
| `attendance.json` | `missed` | Количество пропусков за день |
| `schedule_grid.json` | `dayOfWeek` | День недели (текст), косвенно дата |
| `vedomost.json` | `missedTotal`, `missedBad`, `missedExcused` | Пропуски по студенту/дисциплине (агрегаты) |

Дополнительно к «дате» и «посещаемости» для будущего расширения:
- **Оценки** — сейчас в JSON нет; логично хранить по (дата, группа, студент, дисциплина).
- **Замены** — нет отдельного поля; можно добавить флаг/тип занятия в расписании.
- **Итоги за день/неделю** — считаются на бэкенде из `schedule` + `attendance`, не хранятся отдельным файлом.

---

## 4. Структура каждого JSON

### 4.1. students.json

```json
{
  "totalStudents": 1932,
  "departments": [
    {
      "department": "Отделение экономики",
      "groups": [
        {
          "group": "1бд1",
          "students": [
            {
              "serialNumber": 1,
              "numberInGroup": 1,
              "fullName": "ФИО",
              "status": "Студент"
            }
          ]
        }
      ]
    }
  ]
}
```

- **Связи:** нет внешних ключей; сопоставление с другими файлами по строкам `department`, `group`, `fullName` (или `studentName` в schedule).

### 4.2. schedule.json

```json
{
  "period": "16.02.2026 - 22.02.2026",
  "groups": [
    {
      "group": "2д1",
      "department": "Отделение креативных индустрий",
      "students": [
        {
          "studentName": "ФИО",
          "numberInGroup": 0,
          "records": [
            {
              "date": "16.02.2026 0:00:00",
              "lessonNumber": 1,
              "discipline": "Название",
              "teacher": "Фамилия И.О.",
              "attendance": false
            }
          ],
          "totalCount": 19
        }
      ]
    }
  ]
}
```

- **Связи:** `group` + `department` + `studentName` с `students.json` (по ФИО и группе); с посещаемостью — по `department` + `group` + `student` и дате (дата в schedule — `DD.MM.YYYY`, в attendance — `YYYY-MM-DD`, нужна нормализация).

### 4.3. schedule_grid.json

```json
{
  "period": "16.02.2026 - 22.02.2026",
  "weekStartDate": "2026-02-16",
  "records": [
    {
      "dayOfWeek": "Понедельник",
      "lessonNumber": 1,
      "group": "3ис1",
      "discipline": "Название",
      "teacher": "Фамилия И.О.",
      "location": "Корпус-Аудитория 407"
    }
  ]
}
```

- **Связи:** `group` с остальными файлами; дата восстанавливается по `weekStartDate` + `dayOfWeek`.

### 4.4. attendance.json

```json
[
  {
    "department": "Отделение ...",
    "groups": [
      {
        "group": "1ис1",
        "students": [
          {
            "student": "ФИО",
            "attendance": [
              { "date": "2026-01-12", "missed": 8 },
              { "date": "2026-01-15", "missed": 4 }
            ]
          }
        ]
      }
    ]
  }
]
```

- **Связи:** `department` + `group` + `student` (ФИО); дата в формате `YYYY-MM-DD`. Для сверки с расписанием дату приводят к одному формату.

### 4.5. vedomost.json

```json
[
  {
    "department": "Отделение ...",
    "totalMissed": 92,
    "specialties": [
      {
        "specialty": "09.02.07 ...",
        "totalMissed": 84,
        "groups": [
          {
            "group": "1ис1",
            "totalMissed": 12,
            "students": [
              {
                "student": "Дисциплина или ФИО",
                "missedTotal": 6,
                "missedBad": 0,
                "missedExcused": 6
              }
            ]
          }
        ]
      }
    ]
  }
]
```

- **Связи:** по `department`, `group`; поле `student` может быть названием дисциплины или ФИО — структура смешанная.

---

## 5. Связи между файлами (ключи)

| Связь | Ключи | Примечание |
|-------|--------|------------|
| students ↔ schedule | `department`, `group`, `fullName` ↔ `studentName` | Один контингент — много записей расписания |
| students ↔ attendance | `department`, `group`, `fullName` ↔ `student` | Один студент — много дат в attendance |
| schedule ↔ attendance | `department`, `group`, `studentName`/`student`, дата | Сверка: planned из schedule, present = не в attendance с missed>0 за эту дату |
| schedule_grid ↔ schedule | `group`, `dayOfWeek`+`lessonNumber` ↔ date+lessonNumber | Сетка — «шаблон» недели; schedule — детализация по студентам и датам |

Даты: в **schedule** — `DD.MM.YYYY 0:00:00`, в **attendance** — `YYYY-MM-DD`. В коде используется нормализация к `YYYY-MM-DD` для фильтрации.

---

## 6. Использование на бэкенде

- **Reconciliation:** `schedule.json` → кто должен быть (planned) по дате/паре; `attendance.json` → у кого `missed > 0` за дату (absent). Present = planned − absent.
- **Dashboard / drill:** список отделений и групп строится из `attendance.json` (и при необходимости из `students.json` для контингента); агрегаты считаются по плоскому списку `FlatRecord(Department, Group, Student, Date, Missed)`.
- **Исторический дашборд:** те же данные за выбранную дату; таблица «6 пар» — по одной сверке на каждую пару (ReconcileDayLesson).

Моковые данные на фронте убраны: везде используются ответы API, которые строятся из этих JSON (и при отсутствии данных — fallback в виде «---» или нулей, без дублирования одного дня на все пары).
