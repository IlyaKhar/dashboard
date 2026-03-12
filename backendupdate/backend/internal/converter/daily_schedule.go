package converter

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
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
	records, err := parseDailyXLS(inputFile, pythonScript)
	if err != nil {
		return err
	}

	log.Printf("[DailySchedule] Распарсено записей: %d", len(records))

	// Дата = сегодня
	today := time.Now().Format("02.01.2006")
	todayISO := time.Now().Format("2006-01-02")
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
func parseDailyXLS(inputFile, pythonScript string) ([]dailyRecord, error) {
	// Конвертируем XLS → XLSX если нужно
	xlsxFile := inputFile
	if strings.HasSuffix(strings.ToLower(inputFile), ".xls") && !strings.HasSuffix(strings.ToLower(inputFile), ".xlsx") {
		xlsxFile = strings.TrimSuffix(inputFile, filepath.Ext(inputFile)) + ".xlsx"
		if _, err := os.Stat(xlsxFile); os.IsNotExist(err) {
			if pythonScript == "" {
				pythonScript = filepath.Join(filepath.Dir(inputFile), "xls_to_xlsx.py")
			}
			if err := convertXLSToXLSX(inputFile, xlsxFile, pythonScript); err != nil {
				return nil, fmt.Errorf("ошибка конвертации XLS → XLSX: %v", err)
			}
		}
	}

	f, err := excelize.OpenFile(xlsxFile)
	if err != nil {
		return nil, fmt.Errorf("ошибка открытия файла: %v", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return nil, fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("ошибка чтения строк: %v", err)
	}

	log.Printf("[DailySchedule] Прочитано строк: %d", len(rows))

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
		return nil, fmt.Errorf("не найдена строка с заголовками групп")
	}

	// Извлекаем список групп
	headerRowData := rows[headerRow]
	groupCols := make(map[string]int)
	for j := 1; j < len(headerRowData); j++ {
		cellValue := strings.TrimSpace(headerRowData[j])
		if cellValue != "" && isGroupCode(cellValue) {
			groupCols[cellValue] = j
		}
	}

	log.Printf("[DailySchedule] Найдено групп: %d", len(groupCols))

	// Парсим данные
	var records []dailyRecord
	dayNames := map[string]bool{
		"понедельник": true, "вторник": true, "среда": true,
		"четверг": true, "пятница": true, "суббота": true, "воскресенье": true,
	}

	dataStartRow := headerRow + 1
	var currentLessonNumber int

	for i := dataStartRow; i < len(rows); i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		firstCell := ""
		if len(row) > 0 {
			firstCell = strings.TrimSpace(row[0])
		}
		lessonNumCell := ""
		if len(row) > 1 {
			lessonNumCell = strings.TrimSpace(row[1])
		}

		// Строка с днём недели
		firstLower := strings.ToLower(firstCell)
		if dayNames[firstLower] {
			if num := parseLessonNumber(lessonNumCell); num > 0 {
				currentLessonNumber = num
				records = append(records, extractRecordsFromRow(row, groupCols, currentLessonNumber)...)
			}
			continue
		}

		// Строка без дня недели, но с номером пары
		if firstCell == "" && lessonNumCell != "" {
			if num := parseLessonNumber(lessonNumCell); num > 0 {
				currentLessonNumber = num
				records = append(records, extractRecordsFromRow(row, groupCols, currentLessonNumber)...)
			}
		}
	}

	return records, nil
}

// extractRecordsFromRow извлекает записи из одной строки для всех групп
func extractRecordsFromRow(row []string, groupCols map[string]int, lessonNumber int) []dailyRecord {
	var records []dailyRecord
	for group, colIdx := range groupCols {
		if colIdx >= len(row) {
			continue
		}
		cellValue := strings.TrimSpace(row[colIdx])
		if cellValue == "" {
			continue
		}
		discipline, teacher, _ := parseLessonCell(cellValue)
		if discipline != "" {
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
