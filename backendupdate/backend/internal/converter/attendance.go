package converter

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Типы данных для посещаемости

type AttendanceRecord struct {
	Date         string `json:"date"`
	Missed       int    `json:"missed"`
	LessonNumber int    `json:"lessonNumber,omitempty"`
	Discipline   string `json:"discipline,omitempty"`
}

type Student struct {
	Student    string             `json:"student"`
	Attendance []AttendanceRecord `json:"attendance"`
}

type Group struct {
	Group    string    `json:"group"`
	Students []Student `json:"students"`
}

type Department struct {
	Department string  `json:"department"`
	Groups     []Group `json:"groups"`
}

// readAttendanceRecords читает файл attendance*.json и разворачивает его
// в плоский список записей посещаемости (department, group, student, date, missed).
func readAttendanceRecords(path string) ([]attendanceRecordItem, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}

	var departments []Department
	if err := json.Unmarshal(data, &departments); err != nil {
		return nil, fmt.Errorf("ошибка парсинга JSON посещаемости: %w", err)
	}

	out := make([]attendanceRecordItem, 0, 1024)
	for _, d := range departments {
		for _, g := range d.Groups {
			for _, s := range g.Students {
				for _, rec := range s.Attendance {
					if rec.Date == "" || rec.Missed <= 0 {
						continue
					}
					out = append(out, attendanceRecordItem{
						Department:   d.Department,
						Group:        g.Group,
						Student:      s.Student,
						Date:         rec.Date,
						Missed:       rec.Missed,
						LessonNumber: rec.LessonNumber,
						Discipline:   rec.Discipline,
					})
				}
			}
		}
	}

	return out, nil
}

// ConvertAttendance конвертирует файл посещаемости Excel в JSON
// inputFile - путь к файлу Посещаемость.xlsx
// outputFile - путь к выходному JSON файлу
func ConvertAttendance(inputFile, outputFile string) error {
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		return fmt.Errorf("ошибка открытия файла: %v", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("ошибка чтения строк: %v", err)
	}

	var currentDepartment string
	var currentGroup string
	var currentStudent string

	var records []map[string]interface{}

	for rowIdx, row := range rows {
		if len(row) == 0 {
			continue
		}

		cellName := fmt.Sprintf("A%d", rowIdx+1)
		firstCellValue, err := f.GetCellValue(sheetName, cellName)
		if err != nil {
			if len(row) > 0 {
				firstCellValue = row[0]
			} else {
				firstCellValue = ""
			}
		}
		firstCell := strings.TrimSpace(firstCellValue)

		hoursCellName := fmt.Sprintf("F%d", rowIdx+1)
		hoursValue := 0.0
		hasHours := false

		hoursNumStr, err := f.GetCellValue(sheetName, hoursCellName)
		if err == nil && hoursNumStr != "" {
			if val, err := strconv.ParseFloat(strings.TrimSpace(hoursNumStr), 64); err == nil && val > 0 {
				hoursValue = val
				hasHours = true
			}
		}

		if !hasHours && len(row) > 5 && strings.TrimSpace(row[5]) != "" {
			if val, err := strconv.ParseFloat(strings.TrimSpace(row[5]), 64); err == nil && val > 0 {
				hoursValue = val
				hasHours = true
			}
		}

		dateStr := ""
		if firstCell != "" {
			cellValue, err := f.GetCellValue(sheetName, cellName)
			if err == nil {
				dateStr = parseDateValue(cellValue)
			}
			if dateStr == "" {
				dateStr = parseDateValue(firstCell)
			}
		}

		if dateStr != "" && hasHours {
			if currentDepartment != "" && currentGroup != "" && currentStudent != "" {
				records = append(records, map[string]interface{}{
					"department":   currentDepartment,
					"group":        currentGroup,
					"student":      currentStudent,
					"date":         dateStr,
					"missed":       int(hoursValue),
					"lessonNumber": parseLessonNumberFromRow(row),
				})
			}
			continue
		}

		if firstCell != "" {
			if strings.HasPrefix(firstCell, "Отделение") {
				currentDepartment = firstCell
				currentGroup = ""
				currentStudent = ""
			} else if len(firstCell) <= 10 && len(firstCell) > 0 {
				if firstCell[0] >= '0' && firstCell[0] <= '9' {
					currentGroup = strings.ToLower(firstCell)
					currentStudent = ""
				}
			} else {
				parts := strings.Fields(firstCell)
				if len(parts) == 3 {
					currentStudent = firstCell
				}
			}
		}
	}

	departmentsMap := make(map[string]*Department)

	for _, r := range records {
		dep := r["department"].(string)
		grp := r["group"].(string)
		stu := r["student"].(string)
		date := r["date"].(string)
		missed := r["missed"].(int)
		lessonNumber, _ := r["lessonNumber"].(int)

		dept, exists := departmentsMap[dep]
		if !exists {
			dept = &Department{
				Department: dep,
				Groups:     []Group{},
			}
			departmentsMap[dep] = dept
		}

		var groupObj *Group
		for i := range dept.Groups {
			if dept.Groups[i].Group == grp {
				groupObj = &dept.Groups[i]
				break
			}
		}
		if groupObj == nil {
			dept.Groups = append(dept.Groups, Group{
				Group:    grp,
				Students: []Student{},
			})
			groupObj = &dept.Groups[len(dept.Groups)-1]
		}
		var studentObj *Student
		for i := range groupObj.Students {
			if groupObj.Students[i].Student == stu {
				studentObj = &groupObj.Students[i]
				break
			}
		}
		if studentObj == nil {
			groupObj.Students = append(groupObj.Students, Student{
				Student:    stu,
				Attendance: []AttendanceRecord{},
			})
			studentObj = &groupObj.Students[len(groupObj.Students)-1]
		}

		studentObj.Attendance = append(studentObj.Attendance, AttendanceRecord{
			Date:         date,
			Missed:       missed,
			LessonNumber: lessonNumber,
			Discipline:   "",
		})
	}

	departments := make([]Department, 0, len(departmentsMap))
	for _, dept := range departmentsMap {
		departments = append(departments, *dept)
	}

	outputPath, err := filepath.Abs(outputFile)
	if err != nil {
		return fmt.Errorf("ошибка получения пути: %v", err)
	}

	jsonData, err := json.MarshalIndent(departments, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка серилизации JSON: %v", err)
	}

	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %v", err)
	}

	fmt.Printf(" Конвертация посещаемости завершена. Отделений: %d\n", len(departments))
	fmt.Printf("   Файл сохранён: %s\n", outputPath)
	return nil
}

// UpdateAttendanceHistory обновляет накопительный файл истории посещаемости.
// newFile     — свежий attendance.json за один день;
// historyFile — накопительный attendance_history.json (создаётся, если не существует).
func UpdateAttendanceHistory(newFile, historyFile string) error {
	// Читаем новую порцию данных
	newRecords, err := readAttendanceRecords(newFile)
	if err != nil {
		return fmt.Errorf("не удалось прочитать новый attendance.json: %w", err)
	}
	if len(newRecords) == 0 {
		// Нечего добавлять
		return nil
	}

	// Читаем старую историю, если есть
	historyRecords := []attendanceRecordItem{}
	if _, err := os.Stat(historyFile); err == nil {
		if recs, err := readAttendanceRecords(historyFile); err == nil {
			historyRecords = recs
		}
	}

	// Собираем map по ключу dept|group|student|date → missed.
	type key struct {
		dept, group, student, date string
		lessonNumber               int
		discipline                 string
	}
	m := make(map[key]int)

	for _, r := range historyRecords {
		if r.Missed <= 0 || r.Date == "" {
			continue
		}
		k := key{
			dept: r.Department, group: r.Group, student: r.Student, date: r.Date,
			lessonNumber: r.LessonNumber, discipline: r.Discipline,
		}
		// В истории уже агрегировано, просто переносим значение.
		m[k] = r.Missed
	}

	// Новые данные имеют приоритет: обновляем/добавляем записи за тот же день.
	for _, r := range newRecords {
		if r.Missed <= 0 || r.Date == "" {
			continue
		}
		k := key{
			dept: r.Department, group: r.Group, student: r.Student, date: r.Date,
			lessonNumber: r.LessonNumber, discipline: r.Discipline,
		}
		m[k] = r.Missed
	}

	// Превращаем map обратно в слайс записей и пересобираем JSON через существующую writeAttendanceJSON.
	merged := make([]attendanceRecordItem, 0, len(m))
	for k, missed := range m {
		merged = append(merged, attendanceRecordItem{
			Department:   k.dept,
			Group:        k.group,
			Student:      k.student,
			Date:         k.date,
			Missed:       missed,
			LessonNumber: k.lessonNumber,
			Discipline:   k.discipline,
		})
	}

	if err := writeAttendanceJSON(merged, historyFile); err != nil {
		return fmt.Errorf("ошибка записи истории посещаемости: %w", err)
	}

	return nil
}

func parseDateValue(value string) string {
	value = strings.TrimSpace(value)
	if value == "" {
		return ""
	}

	// Если пришло число — это может быть сериал Excel, но
	// пропущенные часы (2, 4, 6, 8 ...) нам тоже приходят как числа.
	// Для безопасности считаем датой только большие значения (≈ после 1960‑х),
	// чтобы не путать часы с датой.
	if num, err := strconv.ParseFloat(value, 64); err == nil {
		if num >= 30000 && num < 100000 {
			excelEpoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
			days := int(num)
			date := excelEpoch.AddDate(0, 0, days)
			return date.Format("2006-01-02")
		}
	}

	// Даты с названием месяца, например: "2 апреля 2026", "02.апреля.2026", "3 апр. 26"
	monthMap := map[string]time.Month{
		"январь": time.January, "января": time.January, "янв": time.January,
		"февраль": time.February, "февраля": time.February, "фев": time.February,
		"март": time.March, "марта": time.March, "мар": time.March,
		"апрель": time.April, "апреля": time.April, "апр": time.April,
		"май": time.May, "мая": time.May,
		"июнь": time.June, "июня": time.June, "июн": time.June,
		"июль": time.July, "июля": time.July, "июл": time.July,
		"август": time.August, "августа": time.August, "авг": time.August,
		"сентябрь": time.September, "сентября": time.September, "сен": time.September,
		"октябрь": time.October, "октября": time.October, "окт": time.October,
		"ноябрь": time.November, "ноября": time.November, "но": time.November, "ноя": time.November,
		"декабрь": time.December, "декабря": time.December, "дек": time.December,
	}

	normalizeMonthToken := func(token string) string {
		token = strings.ToLower(strings.TrimSpace(token))
		token = strings.ReplaceAll(token, "ё", "е")
		lettersRe := regexp.MustCompile(`[^а-яa-z]+`)
		token = lettersRe.ReplaceAllString(token, "")
		return token
	}

	monthRe := regexp.MustCompile(`(\d{1,2})\s*[.\-/ ]?\s*([А-Яа-яЁё]+)\s*[.\-/ ]?(?:([0-9]{2,4}))?`)
	if m := monthRe.FindStringSubmatch(value); len(m) >= 3 {
		dd, err := strconv.Atoi(m[1])
		if err == nil && dd >= 1 && dd <= 31 {
			monthToken := normalizeMonthToken(m[2])
			parsedMonth, ok := monthMap[monthToken]
			if ok {
				year := time.Now().Year()
				if len(m) >= 4 && m[3] != "" {
					yy := m[3]
					if len(yy) == 2 {
						year, _ = strconv.Atoi("20" + yy)
					} else {
						year, _ = strconv.Atoi(yy)
					}
				}

				parsed := time.Date(year, parsedMonth, dd, 0, 0, 0, 0, time.UTC)
				return parsed.Format("2006-01-02")
			}
		}
	}

	formats := []string{
		"02.01.2006 15:04:05",
		"02.01.2006 0:00:00",
		"2.1.2006 15:04:05",
		"2.1.2006 0:00:00",
		"02.01.2006",
		"02/01/2006",
		"2006-01-02",
		"02.01.06",
		"02/01/06",
		"2.1.2006",
		"2/1/2006",
		"2.1.06",
		"2/1/06",
		"01/02/2006",
		"01-02-2006",
	}

	for _, format := range formats {
		if parsed, err := time.Parse(format, value); err == nil {
			return parsed.Format("2006-01-02")
		}
	}
	return ""
}

// parseAttendanceRow пытается найти дату и пропущенные часы в строке.
// Дату берём только из "датоподобных" ячеек (первая колонка),
// а если её там нет — используем defaultDate (например, из периода сводной ведомости).
func parseAttendanceRow(row []string, defaultDate string) (date string, missed int, lessonNumber int) {
	// 1. Пытаемся вытащить дату из первой ячейки
	if len(row) > 0 {
		if d := parseDateValue(row[0]); d != "" {
			date = d
		}
	}

	// 2. Ищем пропущенные часы (колонка F и правее — последние числовые значения)
	for i := len(row) - 1; i >= 0 && i >= 5; i-- {
		if num, err := strconv.ParseFloat(strings.TrimSpace(row[i]), 64); err == nil && num > 0 {
			missed = int(num)
			break
		}
	}

	// 3. Если в строке даты нет, но есть defaultDate из периода — подставляем её
	if date == "" {
		date = defaultDate
	}

	lessonNumber = parseLessonNumberFromRow(row)
	return date, missed, lessonNumber
}

// parseLessonNumberFromRow пытается вытащить номер пары из начала строки.
// Возвращает 0, если пара не определена (для обратной совместимости со старым форматом).
func parseLessonNumberFromRow(row []string) int {
	limit := len(row)
	if limit > 4 {
		limit = 4
	}

	// В ведомости номер пары часто идёт просто как "1..6" в первой колонке
	// (без слова "пара"), поэтому сначала обрабатываем такой формат.
	if len(row) > 0 {
		first := strings.TrimSpace(row[0])
		if first == "1" || first == "2" || first == "3" || first == "4" || first == "5" || first == "6" {
			n, _ := strconv.Atoi(first)
			return n
		}
	}

	// Затем поддерживаем текстовые форматы с явным словом "пара".
	for i := 0; i < limit; i++ {
		cell := strings.TrimSpace(row[i])
		if cell == "" {
			continue
		}
		low := strings.ToLower(cell)
		if !strings.Contains(low, "пара") {
			continue
		}

		// Берём первую цифру 1..6 в тексте с "пара".
		for _, r := range low {
			if r >= '1' && r <= '6' {
				return int(r - '0')
			}
		}
	}

	return 0
}

func extractLeadingLessonNumber(value string) int {
	runes := []rune(value)
	digits := make([]rune, 0, 2)
	for _, r := range runes {
		if r >= '0' && r <= '9' {
			digits = append(digits, r)
			if len(digits) >= 2 {
				break
			}
			continue
		}
		if len(digits) > 0 {
			break
		}
	}
	if len(digits) == 0 {
		return 0
	}
	n, err := strconv.Atoi(string(digits))
	if err != nil {
		return 0
	}
	return n
}
