# Мастер-конвертер ведомости

Универсальный конвертер: один файл `ведомость.xls` → все JSON файлы.

## Использование

### CLI

```bash
go run cmd/convert-master/main.go -in путь/к/ведомость.xls -out путь/к/выходной/директории
```

Параметры:
- `-in` - путь к файлу ведомость.xls/.xlsx (по умолчанию из конфигурации)
- `-out` - директория для сохранения JSON (по умолчанию `public/`)

### HTTP API

```bash
POST /api/admin/convert/master
Content-Type: multipart/form-data

file: ведомость.xls
```

Требуется авторизация с ролью `admin`.

## Что делает

1. Конвертирует `ведомость.xls` → `ведомость.xlsx` (через Python скрипт)
2. Анализирует структуру файла
3. Извлекает данные:
   - **Контингент студентов** → `students.json`
   - **Детальная посещаемость по датам** → `attendance.json`
   - **Сводная ведомость пропусков** → `vedomost.json`
4. Сохраняет все JSON файлы в указанную директорию

## Ответ API

```json
{
  "ok": true,
  "message": "Мастер-конвертация завершена",
  "outputs": {
    "students": "/path/to/students.json",
    "attendance": "/path/to/attendance.json",
    "vedomost": "/path/to/vedomost.json"
  },
  "warnings": ["Не найдены данные контингента студентов"],
  "errors": []
}
```
