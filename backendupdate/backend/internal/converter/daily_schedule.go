package converter

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// dailyRecord — запись о паре для одной группы
type dailyRecord struct {
	Group        string
	LessonNumber int
	Discipline   string
	Teacher      string
}

// parseDayAndOptionalDate пытается распознать строку дня недели, где рядом может быть дата.
// Примеры:
// - "Понедельник"
// - "Понедельник 17.03.2026"
// - "Понедельник 17.03.26"
// - "Понедельник 17.03"
func parseDayAndOptionalDate(s string) (isDay bool, dateDMY string, dateISO string) {
	s = strings.TrimSpace(s)
	if s == "" {
		return false, "", ""
	}

	// Мягко вытаскиваем: "день" + "дата" (если есть)
	// Важно: не используем `\b` (word boundary) — для кириллицы он иногда работает нестабильно.
	dayRe := regexp.MustCompile(`(?i)(понедельник|вторник|среда|четверг|пятница|суббота|воскресенье)`)
	if !dayRe.MatchString(s) {
		return false, "", ""
	}
	isDay = true

	normalizeMonthToken := func(token string) string {
		token = strings.ToLower(strings.TrimSpace(token))
		token = strings.ReplaceAll(token, "ё", "е")
		// оставляем только буквы
		lettersRe := regexp.MustCompile(`[^а-яa-z]+`)
		token = lettersRe.ReplaceAllString(token, "")
		return token
	}

	monthMap := map[string]time.Month{
		// январь
		"январь":  time.January, "января": time.January, "янв": time.January,
		// февраль
		"февраль": time.February, "февраля": time.February, "фев": time.February,
		// март
		"март": time.March, "марта": time.March, "мар": time.March,
		// апрель
		"апрель": time.April, "апреля": time.April, "апр": time.April,
		// май
		"май": time.May, "мая": time.May,
		// июнь
		"июнь": time.June, "июня": time.June, "июн": time.June,
		// июль
		"июль": time.July, "июля": time.July, "июл": time.July,
		// август
		"август": time.August, "августа": time.August, "авг": time.August,
		// сентябрь
		"сентябрь": time.September, "сентября": time.September, "сен": time.September,
		// октябрь
		"октябрь": time.October, "октября": time.October, "окт": time.October,
		// ноябрь
		"ноябрь": time.November, "ноября": time.November, "ноя": time.November, "но": time.November,
		// декабрь
		"декабрь": time.December, "декабря": time.December, "дек": time.December,
	}

	// 1) Сначала пытаемся распарсить числовой вариант: "02.04.2026", "2/4/26"
	// Ищем дату где угодно в строке.
	dateRe := regexp.MustCompile(`(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?`)
	m := dateRe.FindStringSubmatch(s)
	if len(m) >= 3 {
		dd, err1 := strconv.Atoi(m[1])
		mm, err2 := strconv.Atoi(m[2])
		yy := ""
		if len(m) >= 4 {
			yy = m[3]
		}

		if err1 == nil && err2 == nil && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31 {
			year := time.Now().Year()
			if yy != "" {
				if len(yy) == 2 {
					year, _ = strconv.Atoi("20" + yy)
				} else {
					year, _ = strconv.Atoi(yy)
				}
			}
			parsed := time.Date(year, time.Month(mm), dd, 0, 0, 0, 0, time.UTC)
			return true, parsed.Format("02.01.2006"), parsed.Format("2006-01-02")
		}
	}

	// 2) Затем пытаемся распарсить вариант с названием месяца: "2 апреля 2026", "02.апреля.2026"
	// Пробуем матчить: день + слово-месяц + (опционально) год.
	monthRe := regexp.MustCompile(`(\d{1,2})\s*[.\-/ ]?\s*([А-Яа-яЁё]+)\s*[.\-/ ]?(?:([0-9]{2,4}))?`)
	m2 := monthRe.FindStringSubmatch(s)
	if len(m2) >= 3 {
		dd, err := strconv.Atoi(m2[1])
		if err == nil && dd >= 1 && dd <= 31 {
			monthToken := normalizeMonthToken(m2[2])
			parsedMonth, ok := monthMap[monthToken]
			if ok {
				year := time.Now().Year()
				if len(m2) >= 4 && m2[3] != "" {
					yy := m2[3]
					if len(yy) == 2 {
						year, _ = strconv.Atoi("20" + yy)
					} else {
						year, _ = strconv.Atoi(yy)
					}
				}
				parsed := time.Date(year, parsedMonth, dd, 0, 0, 0, 0, time.UTC)
				return true, parsed.Format("02.01.2006"), parsed.Format("2006-01-02")
			}
		}
	}

	return true, "", ""
}

// isLikelyDateOrDayCell — ячейка похожа на «день недели» или дату, а не на номер пары.
// Иначе parseLessonNumber("02.04.2026") даст 02 → вторая пара вместо первой.
func isLikelyDateOrDayCell(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	low := strings.ToLower(s)
	for _, p := range []string{
		"понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье",
	} {
		if strings.HasPrefix(low, p) {
			return true
		}
	}
	// Полная дата DD.MM.YYYY в строке (как в объединённой ячейке дня).
	dateInString := regexp.MustCompile(`\d{1,2}\s*\.\s*\d{1,2}\s*\.\s*\d{2,4}`)
	return dateInString.MatchString(s)
}

// parseLessonNumberCell — номер пары из колонки «Номер пары» (1–6), без подстановки колонки дня.
func parseLessonNumberCell(s string) int {
	s = strings.TrimSpace(s)
	if s == "" || isLikelyDateOrDayCell(s) {
		return 0
	}
	// Берём номер пары по первой цифре 1..6 в начале строки.
	// Это ловит кейсы вроде "1к", "1а", "1-я", "№1" где текущий parseLessonNumber()
	// может не сработать из-за отсутствия word-boundary между цифрой и буквой.
	leadingRe := regexp.MustCompile(`^\s*(?:№\s*)?([1-7])\b?`)
	if m := leadingRe.FindStringSubmatch(s); len(m) == 2 {
		if n, err := strconv.Atoi(m[1]); err == nil && n >= 1 && n <= 7 {
			return n
		}
	}
	// Текстовые «1» из Excel — часто одна цифра; берём только явные 1–7.
	if len(s) <= 4 {
		if n, err := strconv.Atoi(strings.TrimSpace(s)); err == nil && n >= 1 && n <= 7 {
			return n
		}
	}
	n := parseLessonNumber(s)
	if n < 1 || n > 7 {
		return 0
	}
	return n
}

// ConvertDailySchedule парсит файл РасписаниеНаДату.xls (расписание на 1 день)
// и создаёт два файла:
//   - schedule.json — расписание только на сегодня (оперативный режим)
//   - schedule_history.json — накопленная история за все дни (исторический режим)
//
// Логика:
//  1. Парсим xls → извлекаем группы и пары за 1 день
//  2. Дата = сегодня (файл не содержит явной даты)
//  3. Пишем schedule.json = только сегодня (перезаписываем каждый раз)
//  4. Читаем schedule_history.json, удаляем записи за сегодня (если были), добавляем новые, сохраняем
func ConvertDailySchedule(inputFile, scheduleOutputFile, studentsFile, pythonScript string) error {
	log.Printf("[DailySchedule] Обработка файла: %s", inputFile)

	// Парсим файл → получаем записи
	records, parsedDMY, parsedISO, err := parseDailyXLS(inputFile, pythonScript)
	if err != nil {
		return err
	}

	log.Printf("[DailySchedule] Распарсено записей: %d", len(records))

	// Нормализация возможного смещения в исходной выгрузке:
	// часто в "РасписаниеНаДату" пары идут как 2..6 (фактически весь день сдвинут на +1),
	// тогда корректно привести их к 1..5.
	lessonCnt := make(map[int]int, 7)
	lessonMin := 0
	for _, r := range records {
		if r.LessonNumber < 1 || r.LessonNumber > 7 {
			continue
		}
		lessonCnt[r.LessonNumber]++
		if lessonMin == 0 || r.LessonNumber < lessonMin {
			lessonMin = r.LessonNumber
		}
	}
	hasFullTail7 := lessonCnt[2] > 0 && lessonCnt[3] > 0 && lessonCnt[4] > 0 && lessonCnt[5] > 0 && lessonCnt[6] > 0 && lessonCnt[7] > 0
	log.Printf(
		"[DailySchedule] Распределение пар: 1=%d 2=%d 3=%d 4=%d 5=%d 6=%d 7=%d",
		lessonCnt[1], lessonCnt[2], lessonCnt[3], lessonCnt[4], lessonCnt[5], lessonCnt[6], lessonCnt[7],
	)
	// Если в исходной выгрузке весь день размечен как 2..7 без 1,
	// это значит, что последняя (6-я) пара реально уехала в значение 7.
	// Сдвигаем на -1 и получаем корректные 1..6.
	if lessonMin == 2 && lessonCnt[1] == 0 && hasFullTail7 {
		for i := range records {
			if records[i].LessonNumber >= 2 && records[i].LessonNumber <= 7 {
				records[i].LessonNumber -= 1
			}
		}
		log.Printf("[DailySchedule] Инференс: сдвиг номеров пар на -1 (обнаружено 2..7 без 1-й пары)")
	}

	// Дата = из файла (если указана рядом с днём недели), иначе "сегодня".
	// ВАЖНО: UI ожидает "operational date" в логике российского календаря,
	// а сервер может быть в другой timezone. Чтобы не сдвигать дату на +/-1 день,
	// используем fallback по UTC+3 (условно "МСК").
	today := parsedDMY
	todayISO := parsedISO
	if today == "" || todayISO == "" {
		nowOperational := time.Now().UTC().Add(5 * time.Hour)
		today = nowOperational.Format("02.01.2006")
		todayISO = nowOperational.Format("2006-01-02")
	}
	log.Printf("[DailySchedule] Дата расписания: %s", today)

	// Загружаем список студентов
	groupStudents, groupDepartments := loadStudentsMap(studentsFile)

	// === 1. schedule.json — только сегодня (перезаписываем) ===
	todayOutput := buildScheduleForDay(records, today, todayISO, groupStudents, groupDepartments)
	if err := saveScheduleJSON(todayOutput, scheduleOutputFile); err != nil {
		return fmt.Errorf("ошибка записи schedule.json: %v", err)
	}
	log.Printf("[DailySchedule] ✓ schedule.json: только сегодня (%s), групп: %d, студентов: %d",
		today, todayOutput.TotalGroups, todayOutput.TotalStudents)

	// === 2. schedule_history.json — накопление ===
	historyFile := strings.Replace(scheduleOutputFile, "schedule.json", "schedule_history.json", 1)
	if err := mergeIntoHistory(records, today, todayISO, groupStudents, groupDepartments, historyFile); err != nil {
		log.Printf("[DailySchedule] Предупреждение: не удалось обновить schedule_history.json: %v", err)
	} else {
		log.Printf("[DailySchedule] ✓ schedule_history.json обновлён (накопление)")
	}

	return nil
}

// parseDailyXLS парсит РасписаниеНаДату.xls и возвращает список записей
func parseDailyXLS(inputFile, pythonScript string) ([]dailyRecord, string, string, error) {
	// Конвертируем XLS → XLSX если нужно
	xlsxFile := inputFile
	if strings.HasSuffix(strings.ToLower(inputFile), ".xls") && !strings.HasSuffix(strings.ToLower(inputFile), ".xlsx") {
		xlsxFile = strings.TrimSuffix(inputFile, filepath.Ext(inputFile)) + ".xlsx"
		if _, err := os.Stat(xlsxFile); os.IsNotExist(err) {
			if pythonScript == "" {
				pythonScript = filepath.Join(filepath.Dir(inputFile), "xls_to_xlsx.py")
			}
			if err := convertXLSToXLSX(inputFile, xlsxFile, pythonScript); err != nil {
				return nil, "", "", fmt.Errorf("ошибка конвертации XLS → XLSX: %v", err)
			}
		}
	}

	f, err := excelize.OpenFile(xlsxFile)
	if err != nil {
		return nil, "", "", fmt.Errorf("ошибка открытия файла: %v", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return nil, "", "", fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, "", "", fmt.Errorf("ошибка чтения строк: %v", err)
	}

	log.Printf("[DailySchedule] Прочитано строк: %d", len(rows))

	// Пытаемся заранее вытащить "базовую" дату расписания из верхних строк листа.
	// В части выгрузок дата может быть не в ячейке "День недели", а где-то в шапке/служебных блоках.
	// Тогда без этой логики today будет падать в fallback "сегодня", и в истории получатся одинаковые дни.
	baseISO := ""
	{
		limit := 30
		if limit > len(rows) {
			limit = len(rows)
		}
		// В верхних строках могут встречаться служебные даты (например, дата согласования).
		// Берём максимальную найденную дату, чтобы не застревать на "вчера", если в файле есть более свежая.
		for i := 0; i < limit; i++ {
			for _, cell := range rows[i] {
				if iso := parseDateValue(cell); iso != "" {
					if baseISO == "" || iso > baseISO {
						baseISO = iso
					}
				}
			}
		}
	}

	// Ищем строку с заголовками групп
	headerRow := -1
	for i := 0; i < 20 && i < len(rows); i++ {
		row := rows[i]
		if len(row) < 2 {
			continue
		}
		firstCell := strings.TrimSpace(row[0])
		if strings.Contains(firstCell, "День недели") || strings.Contains(firstCell, "Номер пары") {
			headerRow = i
			break
		}
		groupCount := 0
		for j := 1; j < len(row) && j < 200; j++ {
			cellValue := strings.TrimSpace(row[j])
			if cellValue != "" && isGroupCode(cellValue) {
				groupCount++
			}
		}
		if groupCount >= 3 {
			headerRow = i
			break
		}
	}

	if headerRow == -1 {
		return nil, "", "", fmt.Errorf("не найдена строка с заголовками групп")
	}

	// Извлекаем список групп
	headerRowData := rows[headerRow]
	groupCols := make(map[string]int)

	// Детерминированно определяем индексы колонок по заголовкам:
	// - dayColIdx: колонка с "День недели"/"Четверг ..."/днём
	// - lessonNumCol: колонка с "Номер пары"
	dayColIdx := 0
	lessonNumCol := -1
	for j, headerCell := range headerRowData {
		cellValue := strings.ToLower(strings.TrimSpace(headerCell))
		if cellValue == "" {
			continue
		}

		if lessonNumCol == -1 &&
			strings.Contains(cellValue, "номер") &&
			(strings.Contains(cellValue, "пары") || strings.Contains(cellValue, "пар")) {
			lessonNumCol = j
		}

		if strings.Contains(cellValue, "день") {
			dayColIdx = j
		}
	}

	// fallback: если не смогли найти по заголовкам (старые/нестандартные файлы)
	if lessonNumCol < 0 {
		firstHeader := ""
		secondHeader := ""
		if len(headerRowData) > 0 {
			firstHeader = strings.ToLower(strings.TrimSpace(headerRowData[0]))
		}
		if len(headerRowData) > 1 {
			secondHeader = strings.ToLower(strings.TrimSpace(headerRowData[1]))
		}
		switch {
		case strings.Contains(firstHeader, "день") && strings.Contains(secondHeader, "номер"):
			lessonNumCol = 1
		case strings.Contains(firstHeader, "номер"):
			lessonNumCol = 0
		default:
			lessonNumCol = 1
		}
	}

	// Группы можно искать по всему заголовку, без предположений про смещение.
	for j := 0; j < len(headerRowData); j++ {
		cellValue := strings.TrimSpace(headerRowData[j])
		if cellValue != "" && isGroupCode(cellValue) {
			groupCols[cellValue] = j
		}
	}

	log.Printf("[DailySchedule] Найдено групп: %d", len(groupCols))
	log.Printf("[DailySchedule] Детект колонок: dayColIdx=%d, lessonNumCol=%d", dayColIdx, lessonNumCol)

	// Парсим данные
	var records []dailyRecord
	var parsedDMY, parsedISO string
	if baseISO != "" {
		if baseDate, err := time.Parse("2006-01-02", baseISO); err == nil {
			parsedDMY = baseDate.Format("02.01.2006")
			parsedISO = baseISO
		}
	}

	dataStartRow := headerRow + 1
	var currentLessonNumber int
	// Частая проблема файлов "РасписаниеНаДату": номер 1-й пары бывает в объединённой/пустой ячейке,
	// а занятия при этом уже есть в строке. В таком случае считаем первую "содержательную" строку парой №1.
	// Также поддерживаем строки с занятиями, где номер пары пустой (берём текущий номер пары).
	groupColIdxs := make([]int, 0, len(groupCols))
	for _, colIdx := range groupCols {
		groupColIdxs = append(groupColIdxs, colIdx)
	}
	sort.Ints(groupColIdxs)

	for i := dataStartRow; i < len(rows); i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		// Важно: в новых файлах много объединённых/пустых ячеек, и GetRows() может "обрезать" строку.
		// Поэтому ключевые значения читаем напрямую из файла по координатам.
		rowNumber := i + 1 // excelize: строки 1-based

		getCell := func(colIdx int) string {
			if colIdx < 0 {
				return ""
			}
			colName, err := excelize.ColumnNumberToName(colIdx + 1) // 1-based
			if err != nil {
				return ""
			}
			v, err := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", colName, rowNumber))
			if err != nil {
				return ""
			}
			return strings.TrimSpace(v)
		}

		dayCell := getCell(dayColIdx)
		// Номер пары из колонки Excel ("Номер пары") бывает ненадёжным на строках
		// с объединёнными ячейками. Поэтому LessonNumber выставляем строго по
		// порядку строк с реальными занятиями внутри блока дня (см. ниже).

		// Строка с днём недели (в новом формате там же может быть дата)
		if isDay, dmy, iso := parseDayAndOptionalDate(dayCell); isDay {
			// При встрече строки с днём недели начинаем отсчёт пар "с чистого листа".
			currentLessonNumber = 0
			// 1) Если дата явно распознана в строке "день недели" — доверяем ей.
			// 2) Если распознана только "день недели", пересчитываем дату относительно baseISO.
			if dmy != "" && iso != "" {
				parsedDMY, parsedISO = dmy, iso
			} else if parsedISO != "" {
				low := strings.ToLower(strings.TrimSpace(dayCell))
				dayOffset := -1
				switch {
				case strings.Contains(low, "понедельник"):
					dayOffset = 0
				case strings.Contains(low, "вторник"):
					dayOffset = 1
				case strings.Contains(low, "среда"):
					dayOffset = 2
				case strings.Contains(low, "четверг"):
					dayOffset = 3
				case strings.Contains(low, "пятница"):
					dayOffset = 4
				case strings.Contains(low, "суббота"):
					dayOffset = 5
				case strings.Contains(low, "воскресенье"):
					dayOffset = 6
				}

				if dayOffset >= 0 {
					if baseDate, err := time.Parse("2006-01-02", parsedISO); err == nil {
						baseOffset := 0
						switch baseDate.Weekday() {
						case time.Monday:
							baseOffset = 0
						case time.Tuesday:
							baseOffset = 1
						case time.Wednesday:
							baseOffset = 2
						case time.Thursday:
							baseOffset = 3
						case time.Friday:
							baseOffset = 4
						case time.Saturday:
							baseOffset = 5
						case time.Sunday:
							baseOffset = 6
						}

						computed := baseDate.AddDate(0, 0, dayOffset-baseOffset)
						parsedDMY = computed.Format("02.01.2006")
						parsedISO = computed.Format("2006-01-02")
					}
				}
			}

			// Иногда первый урок может лежать прямо в первой строке "день недели".
			if rowHasLessons(f, sheetName, rowNumber, groupColIdxs) {
				currentLessonNumber = 1
				records = append(records, extractRecordsFromRowCells(f, sheetName, rowNumber, groupCols, currentLessonNumber)...)
			}
			continue
		}

		// Не полагаемся на `dayCell` (он может быть заполнен/не заполнен из-за объединённых ячеек).
		// Если в строке реально есть занятия по группам — это очередная пара внутри блока дня.
		if rowHasLessons(f, sheetName, rowNumber, groupColIdxs) {
			if currentLessonNumber == 0 {
				currentLessonNumber = 1
			} else {
				currentLessonNumber++
			}
			records = append(records, extractRecordsFromRowCells(f, sheetName, rowNumber, groupCols, currentLessonNumber)...)
		}
	}

	return records, parsedDMY, parsedISO, nil
}

func rowHasLessons(f *excelize.File, sheetName string, rowNumber int, groupColIdxs []int) bool {
	for _, colIdx := range groupColIdxs {
		colName, err := excelize.ColumnNumberToName(colIdx + 1)
		if err != nil {
			continue
		}
		cellValueRaw, err := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", colName, rowNumber))
		if err != nil {
			continue
		}
		cellValue := strings.TrimSpace(cellValueRaw)
		if cellValue == "" {
			continue
		}
		// Иногда в ячейке занятие помечается только прочерком (дистант),
		// а дисциплина/аудитория могут быть пустыми. Тогда `parseLessonCell`
		// не вернёт discipline, но занятие всё равно есть.
		if cellValue == "-" || cellValue == "—" || cellValue == "–" {
			return true
		}

		discipline, teacher, location := parseLessonCell(cellValue)
		if discipline != "" || teacher != "" || location != "" {
			return true
		}
	}
	return false
}

// extractRecordsFromRowCells извлекает записи из одной строки для всех групп, читая значения ячеек напрямую из файла.
func extractRecordsFromRowCells(
	f *excelize.File,
	sheetName string,
	rowNumber int,
	groupCols map[string]int,
	lessonNumber int,
) []dailyRecord {
	var records []dailyRecord
	for group, colIdx := range groupCols {
		colName, err := excelize.ColumnNumberToName(colIdx + 1)
		if err != nil {
			continue
		}
		cellValueRaw, err := f.GetCellValue(sheetName, fmt.Sprintf("%s%d", colName, rowNumber))
		if err != nil {
			continue
		}
		cellValue := strings.TrimSpace(cellValueRaw)
		if cellValue == "" {
			continue
		}
		discipline, teacher, location := parseLessonCell(cellValue)

		// Для некоторых форматов (часто в 1-й паре) дисциплина может быть пустой,
		// но занятие присутствует (например, дистант через прочерк). В таком случае
		// используем location как суррогат discipline, чтобы запись попала в schedule.json.
		if discipline == "" && location != "" {
			discipline = location
		}

		// Записываем занятие, если хотя бы удалось извлечь смысловую часть.
		if discipline != "" || teacher != "" {
			records = append(records, dailyRecord{
				Group:        group,
				LessonNumber: lessonNumber,
				Discipline:   discipline,
				Teacher:      teacher,
			})
		}
	}
	return records
}

// loadStudentsMap загружает студентов из students.json
func loadStudentsMap(studentsFile string) (groupStudents map[string][]string, groupDepartments map[string]string) {
	groupStudents = make(map[string][]string)
	groupDepartments = make(map[string]string)

	if studentsFile == "" {
		return
	}

	type studentsRoot struct {
		Departments []struct {
			Department string `json:"department"`
			Groups     []struct {
				Group    string `json:"group"`
				Students []struct {
					FullName string `json:"fullName"`
				} `json:"students"`
			} `json:"groups"`
		} `json:"departments"`
	}

	studentsData, err := os.ReadFile(studentsFile)
	if err != nil {
		return
	}
	var students studentsRoot
	if err := json.Unmarshal(studentsData, &students); err != nil {
		return
	}
	for _, dept := range students.Departments {
		for _, grp := range dept.Groups {
			var names []string
			for _, st := range grp.Students {
				if st.FullName != "" {
					names = append(names, st.FullName)
				}
			}
			key := strings.ToLower(grp.Group)
			groupStudents[key] = names
			if dept.Department != "" {
				groupDepartments[key] = dept.Department
			}
		}
	}
	return
}

// buildScheduleForDay создаёт LessonsOutput только за один день
func buildScheduleForDay(records []dailyRecord, today, todayISO string, groupStudents map[string][]string, groupDepartments map[string]string) LessonsOutput {
	dateStr := today + " 0:00:00"

	groupsMap := make(map[string]*GroupLessons)
	var groups []GroupLessons

	for _, rec := range records {
		groupKey := strings.ToLower(rec.Group)

		groupObj, exists := groupsMap[groupKey]
		if !exists {
			deptName := groupDepartments[groupKey]
			if deptName == "" {
				deptName = departmentForGroup(rec.Group)
			}
			newGroup := GroupLessons{
				Group:      rec.Group,
				Department: deptName,
				Students:   []StudentLessons{},
			}
			groups = append(groups, newGroup)
			groupsMap[groupKey] = &groups[len(groups)-1]
			groupObj = groupsMap[groupKey]
		}

		studentsList := groupStudents[groupKey]
		if len(studentsList) == 0 {
			continue
		}

		for _, studentName := range studentsList {
			var studentLessons *StudentLessons
			for j := range groupObj.Students {
				if groupObj.Students[j].StudentName == studentName {
					studentLessons = &groupObj.Students[j]
					break
				}
			}
			if studentLessons == nil {
				groupObj.Students = append(groupObj.Students, StudentLessons{
					StudentName: studentName,
					Records:     []LessonRecord{},
				})
				studentLessons = &groupObj.Students[len(groupObj.Students)-1]
				groupObj.TotalStudents++
			}

			studentLessons.Records = append(studentLessons.Records, LessonRecord{
				Date:         dateStr,
				LessonNumber: rec.LessonNumber,
				Discipline:   rec.Discipline,
				Teacher:      rec.Teacher,
				Attendance:   false,
			})
			studentLessons.TotalCount = len(studentLessons.Records)
		}
	}

	totalStudents := 0
	for i := range groups {
		groups[i].TotalStudents = len(groups[i].Students)
		totalStudents += groups[i].TotalStudents
	}

	return LessonsOutput{
		Period:        today + " - " + today,
		Groups:        groups,
		TotalGroups:   len(groups),
		TotalStudents: totalStudents,
	}
}

// mergeIntoHistory читает schedule_history.json, удаляет записи за сегодня, добавляет новые, сохраняет
func mergeIntoHistory(records []dailyRecord, today, todayISO string, groupStudents map[string][]string, groupDepartments map[string]string, historyFile string) error {
	dateStr := today + " 0:00:00"
	datePrefix := today // "DD.MM.YYYY"

	// Читаем существующую историю
	var existing LessonsOutput
	if data, err := os.ReadFile(historyFile); err == nil {
		if err := json.Unmarshal(data, &existing); err != nil {
			log.Printf("[DailySchedule] Ошибка парсинга schedule_history.json, создаём новый: %v", err)
			existing = LessonsOutput{}
		}
	}

	// Удаляем старые записи за сегодня
	for i := range existing.Groups {
		for j := range existing.Groups[i].Students {
			filtered := make([]LessonRecord, 0, len(existing.Groups[i].Students[j].Records))
			for _, rec := range existing.Groups[i].Students[j].Records {
				if !strings.HasPrefix(rec.Date, datePrefix) {
					filtered = append(filtered, rec)
				}
			}
			existing.Groups[i].Students[j].Records = filtered
			existing.Groups[i].Students[j].TotalCount = len(filtered)
		}
	}

	// Индексируем группы
	existingGroupsMap := make(map[string]*GroupLessons)
	for i := range existing.Groups {
		key := strings.ToLower(existing.Groups[i].Group)
		existingGroupsMap[key] = &existing.Groups[i]
	}

	// Добавляем новые записи за сегодня
	for _, rec := range records {
		groupKey := strings.ToLower(rec.Group)

		groupObj, exists := existingGroupsMap[groupKey]
		if !exists {
			deptName := groupDepartments[groupKey]
			if deptName == "" {
				deptName = departmentForGroup(rec.Group)
			}
			newGroup := GroupLessons{
				Group:      rec.Group,
				Department: deptName,
				Students:   []StudentLessons{},
			}
			existing.Groups = append(existing.Groups, newGroup)
			existingGroupsMap[groupKey] = &existing.Groups[len(existing.Groups)-1]
			groupObj = existingGroupsMap[groupKey]
		}

		studentsList := groupStudents[groupKey]
		if len(studentsList) == 0 {
			continue
		}

		for _, studentName := range studentsList {
			var studentLessons *StudentLessons
			for j := range groupObj.Students {
				if groupObj.Students[j].StudentName == studentName {
					studentLessons = &groupObj.Students[j]
					break
				}
			}
			if studentLessons == nil {
				groupObj.Students = append(groupObj.Students, StudentLessons{
					StudentName: studentName,
					Records:     []LessonRecord{},
				})
				studentLessons = &groupObj.Students[len(groupObj.Students)-1]
				groupObj.TotalStudents++
			}

			studentLessons.Records = append(studentLessons.Records, LessonRecord{
				Date:         dateStr,
				LessonNumber: rec.LessonNumber,
				Discipline:   rec.Discipline,
				Teacher:      rec.Teacher,
				Attendance:   false,
			})
			studentLessons.TotalCount = len(studentLessons.Records)
		}
	}

	// Обновляем счётчики
	totalStudents := 0
	for i := range existing.Groups {
		existing.Groups[i].TotalStudents = len(existing.Groups[i].Students)
		totalStudents += existing.Groups[i].TotalStudents
	}
	existing.TotalGroups = len(existing.Groups)
	existing.TotalStudents = totalStudents

	// Обновляем период
	existing.Period = updatePeriod(existing.Period, todayISO)

	return saveScheduleJSON(existing, historyFile)
}

// saveScheduleJSON сохраняет LessonsOutput в JSON файл
func saveScheduleJSON(output LessonsOutput, filePath string) error {
	outputPath, err := filepath.Abs(filePath)
	if err != nil {
		return fmt.Errorf("ошибка получения пути: %v", err)
	}

	jsonData, err := json.MarshalIndent(output, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %v", err)
	}

	return os.WriteFile(outputPath, jsonData, 0644)
}

// updatePeriod обновляет строку периода, расширяя диапазон дат
func updatePeriod(current, newDateISO string) string {
	newDate, err := time.Parse("2006-01-02", newDateISO)
	if err != nil {
		return current
	}
	newDateStr := newDate.Format("02.01.2006")

	if current == "" {
		return newDateStr + " - " + newDateStr
	}

	re := regexp.MustCompile(`(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})`)
	matches := re.FindStringSubmatch(current)
	if len(matches) != 3 {
		return newDateStr + " - " + newDateStr
	}

	startDate, err1 := time.Parse("02.01.2006", matches[1])
	endDate, err2 := time.Parse("02.01.2006", matches[2])
	if err1 != nil || err2 != nil {
		return newDateStr + " - " + newDateStr
	}

	if newDate.Before(startDate) {
		startDate = newDate
	}
	if newDate.After(endDate) {
		endDate = newDate
	}

	return startDate.Format("02.01.2006") + " - " + endDate.Format("02.01.2006")
}
