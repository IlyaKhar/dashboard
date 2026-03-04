package converter

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ScheduleGridRecord - запись о занятии из сетки расписания
type ScheduleGridRecord struct {
	DayOfWeek   string `json:"dayOfWeek"`   // День недели (Понедельник, Вторник...)
	LessonNumber int   `json:"lessonNumber"` // Номер пары (1, 2, 3...)
	Group       string `json:"group"`        // Группа
	Discipline  string `json:"discipline"`   // Дисциплина
	Teacher     string `json:"teacher"`      // Преподаватель
	Location    string `json:"location"`     // Аудитория/локация
}

// ScheduleGridOutput - итоговая структура для расписания из сетки
type ScheduleGridOutput struct {
	Period        string                `json:"period"`        // Период (например, "02.02.2026 - 08.02.2026")
	WeekStartDate string                `json:"weekStartDate"` // Дата начала недели (YYYY-MM-DD)
	Records       []ScheduleGridRecord  `json:"records"`       // Все записи расписания
	Groups        []string              `json:"groups"`         // Список всех групп
	TotalRecords  int                   `json:"totalRecords"`  // Общее количество записей
}

// ConvertScheduleGrid конвертирует файл "расписание.xls" (сетка расписания) в JSON
// inputFile - путь к файлу расписание.xls
// outputFile - путь к выходному JSON файлу
// weekStartDate - дата начала недели (опционально, если пусто - используется текущая неделя)
func ConvertScheduleGrid(inputFile, outputFile, weekStartDate string) error {
	fmt.Printf("[ScheduleGrid] Открытие файла: %s\n", inputFile)
	
	// Конвертируем XLS в XLSX если нужно
	xlsxFile := inputFile
	if strings.HasSuffix(strings.ToLower(inputFile), ".xls") {
		xlsxFile = strings.TrimSuffix(inputFile, ".xls") + ".xlsx"
		if _, err := os.Stat(xlsxFile); os.IsNotExist(err) {
			// Конвертируем через Python скрипт (используем функцию из statement.go)
			// Ищем скрипт в нескольких местах
			pythonScript := filepath.Join(filepath.Dir(inputFile), "xls_to_xlsx.py")
			if _, err := os.Stat(pythonScript); os.IsNotExist(err) {
				// Пробуем найти в других местах
				possiblePaths := []string{
					filepath.Join(filepath.Dir(filepath.Dir(inputFile)), "converter", "xls_to_xlsx.py"),
					filepath.Join(filepath.Dir(filepath.Dir(filepath.Dir(inputFile))), "backend", "internal", "converter", "xls_to_xlsx.py"),
				}
				for _, path := range possiblePaths {
					if absPath, err := filepath.Abs(path); err == nil {
						if _, err := os.Stat(absPath); err == nil {
							pythonScript = absPath
							break
						}
					}
				}
			}
			// Получаем абсолютный путь
			if absScript, err := filepath.Abs(pythonScript); err == nil {
				pythonScript = absScript
			}
			if absInput, err := filepath.Abs(inputFile); err == nil {
				inputFile = absInput
			}
			if absOutput, err := filepath.Abs(xlsxFile); err == nil {
				xlsxFile = absOutput
			}
			if err := convertXLSToXLSX(inputFile, xlsxFile, pythonScript); err != nil {
				return fmt.Errorf("ошибка конвертации XLS в XLSX: %v", err)
			}
		}
	}

	f, err := excelize.OpenFile(xlsxFile)
	if err != nil {
		return fmt.Errorf("ошибка открытия файла: %v", err)
	}
	defer f.Close()
	fmt.Printf("[ScheduleGrid] Файл успешно открыт\n")

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("ошибка чтения строк: %v", err)
	}
	fmt.Printf("[ScheduleGrid] Прочитано строк: %d\n", len(rows))

	// Определяем структуру таблицы
	// Ищем строку с заголовками групп (обычно первая строка с данными)
	headerRow := -1
	
	// Сначала ищем строку с группами в первых 20 строках
	for i := 0; i < 20 && i < len(rows); i++ {
		row := rows[i]
		if len(row) < 2 {
			continue
		}
		
		// Ищем строку, где первая колонка содержит "День недели" или "Номер пары"
		firstCell := strings.TrimSpace(row[0])
		if strings.Contains(firstCell, "День недели") || strings.Contains(firstCell, "Номер пары") {
			headerRow = i
			fmt.Printf("[ScheduleGrid] Найдена строка с заголовками (День недели/Номер пары) в строке %d\n", i+1)
			break
		}
		
		// Ищем строку с группами (коды типа "1а1", "1бб1" и т.д.)
		groupCount := 0
		for j := 1; j < len(row) && j < 20; j++ {
			cellValue := strings.TrimSpace(row[j])
			if cellValue != "" && isGroupCode(cellValue) {
				groupCount++
			}
		}
		// Если нашли хотя бы 3 группы в строке - это заголовки
		if groupCount >= 3 {
			headerRow = i
			fmt.Printf("[ScheduleGrid] Найдена строка с группами (%d групп) в строке %d\n", groupCount, i+1)
			break
		}
	}

	if headerRow == -1 {
		// Если не нашли, выводим первые строки для отладки
		fmt.Printf("[ScheduleGrid] Отладка: первые 5 строк файла:\n")
		for i := 0; i < 5 && i < len(rows); i++ {
			row := rows[i]
			fmt.Printf("  Строка %d: ", i+1)
			for j := 0; j < len(row) && j < 10; j++ {
				val := strings.TrimSpace(row[j])
				if len(val) > 15 {
					val = val[:15] + "..."
				}
				fmt.Printf("[%d]'%s' ", j, val)
			}
			fmt.Println()
		}
		return fmt.Errorf("не найдена строка с заголовками групп")
	}

	fmt.Printf("[ScheduleGrid] Заголовки найдены в строке %d\n", headerRow+1)

	// Извлекаем список групп из заголовочной строки
	headerRowData := rows[headerRow]
	groups := []string{}
	groupCols := make(map[string]int) // маппинг группы -> индекс колонки
	
	// Первая колонка обычно "День недели" или пустая, пропускаем
	for j := 1; j < len(headerRowData); j++ {
		cellValue := strings.TrimSpace(headerRowData[j])
		if cellValue == "" {
			continue
		}
		// Проверяем, является ли это кодом группы
		if isGroupCode(cellValue) {
			groups = append(groups, cellValue)
			groupCols[cellValue] = j
		}
	}

	fmt.Printf("[ScheduleGrid] Найдено групп: %d\n", len(groups))

	// Определяем начало данных (обычно после заголовков)
	dataStartRow := headerRow + 1
	if dataStartRow >= len(rows) {
		return fmt.Errorf("нет данных после заголовков")
	}
	
	// Парсим расписание
	records := []ScheduleGridRecord{}
	var currentDayOfWeek string
	var currentLessonNumber int

	// Маппинг дней недели
	dayNames := map[string]string{
		"понедельник": "Понедельник",
		"вторник":     "Вторник",
		"среда":       "Среда",
		"четверг":     "Четверг",
		"пятница":     "Пятница",
		"суббота":     "Суббота",
		"воскресенье": "Воскресенье",
	}

	// Обрабатываем строки данных
	for i := dataStartRow; i < len(rows); i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		firstCell := ""
		if len(row) > 0 {
			firstCell = strings.TrimSpace(row[0])
		}
		
		// Номер пары находится в колонке 1 (вторая колонка)
		lessonNumCell := ""
		if len(row) > 1 {
			lessonNumCell = strings.TrimSpace(row[1])
		}

		// Проверяем, является ли первая колонка днём недели
		firstLower := strings.ToLower(firstCell)
		if day, ok := dayNames[firstLower]; ok {
			currentDayOfWeek = day
			currentLessonNumber = 0
			// Сразу проверяем номер пары в этой же строке (колонка 1)
			if num := parseLessonNumber(lessonNumCell); num > 0 {
				currentLessonNumber = num
				// Обрабатываем группы в этой строке (начиная с колонки 2)
				rowRecords := 0
				for group, colIdx := range groupCols {
					if colIdx >= len(row) {
						continue
					}
					cellValue := strings.TrimSpace(row[colIdx])
					if cellValue == "" {
						continue
					}
					discipline, teacher, location := parseLessonCell(cellValue)
					if discipline != "" {
						records = append(records, ScheduleGridRecord{
							DayOfWeek:    currentDayOfWeek,
							LessonNumber: currentLessonNumber,
							Group:        group,
							Discipline:   discipline,
							Teacher:      teacher,
							Location:     location,
						})
						rowRecords++
					}
				}
			}
			continue
		}

		// Если первая колонка пустая, проверяем номер пары в колонке 1
		if firstCell == "" && lessonNumCell != "" {
			if num := parseLessonNumber(lessonNumCell); num > 0 {
				currentLessonNumber = num
				// Обрабатываем все группы в этой строке
				rowRecords := 0
				for group, colIdx := range groupCols {
					if colIdx >= len(row) {
						continue
					}
					cellValue := strings.TrimSpace(row[colIdx])
					if cellValue == "" {
						continue
					}

					// Парсим содержимое ячейки (формат: "1. [Дисциплина] [Преподаватель] [Аудитория]")
					discipline, teacher, location := parseLessonCell(cellValue)
					if discipline != "" {
						records = append(records, ScheduleGridRecord{
							DayOfWeek:    currentDayOfWeek,
							LessonNumber: currentLessonNumber,
							Group:        group,
							Discipline:   discipline,
							Teacher:      teacher,
							Location:     location,
						})
						rowRecords++
					}
				}
				continue
			}
		}

		// Если есть текущий день и номер пары, но строка не обработана выше,
		// возможно это продолжение или данные в других колонках
		if currentDayOfWeek != "" && currentLessonNumber > 0 {
			// Проверяем, есть ли данные в колонках групп
			hasData := false
			for _, colIdx := range groupCols {
				if colIdx < len(row) && strings.TrimSpace(row[colIdx]) != "" {
					hasData = true
					break
				}
			}
			if hasData {
				for group, colIdx := range groupCols {
					if colIdx >= len(row) {
						continue
					}
					cellValue := strings.TrimSpace(row[colIdx])
					if cellValue == "" {
						continue
					}
					discipline, teacher, location := parseLessonCell(cellValue)
					if discipline != "" {
						records = append(records, ScheduleGridRecord{
							DayOfWeek:    currentDayOfWeek,
							LessonNumber: currentLessonNumber,
							Group:        group,
							Discipline:   discipline,
							Teacher:      teacher,
							Location:     location,
						})
					}
				}
			}
		}
	}

	fmt.Printf("[ScheduleGrid] Найдено записей: %d\n", len(records))

	// Определяем период (дату начала недели)
	var startDate time.Time
	if weekStartDate != "" {
		parsed, err := time.Parse("2006-01-02", weekStartDate)
		if err == nil {
			startDate = parsed
		}
	}
	if startDate.IsZero() {
		// Используем начало текущей недели (понедельник)
		now := time.Now()
		weekday := int(now.Weekday())
		if weekday == 0 {
			weekday = 7 // Воскресенье = 7
		}
		startDate = now.AddDate(0, 0, -weekday+1) // Понедельник
	}

	period := fmt.Sprintf("%s - %s",
		startDate.Format("02.01.2006"),
		startDate.AddDate(0, 0, 6).Format("02.01.2006"))

	// Формируем итоговую структуру
	output := ScheduleGridOutput{
		Period:       period,
		WeekStartDate: startDate.Format("2006-01-02"),
		Records:      records,
		Groups:       groups,
		TotalRecords: len(records),
	}

	outputPath, err := filepath.Abs(outputFile)
	if err != nil {
		return fmt.Errorf("ошибка получения пути: %v", err)
	}

	jsonData, err := json.MarshalIndent(output, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %v", err)
	}

	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %v", err)
	}

	fmt.Printf(" Конвертация расписания (сетка) завершена.\n")
	fmt.Printf("   Период: %s\n", period)
	fmt.Printf("   Групп: %d\n", len(groups))
	fmt.Printf("   Записей: %d\n", len(records))
	fmt.Printf("   Файл сохранён: %s\n", outputPath)

	return nil
}

// ConvertScheduleGridToLessonsFormat преобразует сетку расписания в формат lessons.json
// Использует students.json для получения списка студентов по группам
func ConvertScheduleGridToLessonsFormat(gridFile, studentsFile, outputFile, weekStartDate string) error {
	// Загружаем сетку расписания
	type gridRoot struct {
		WeekStartDate string                `json:"weekStartDate"`
		Records       []ScheduleGridRecord   `json:"records"`
	}
	
	gridData, err := os.ReadFile(gridFile)
	if err != nil {
		return fmt.Errorf("ошибка чтения файла сетки: %v", err)
	}
	
	var grid ScheduleGridOutput
	if err := json.Unmarshal(gridData, &grid); err != nil {
		return fmt.Errorf("ошибка парсинга JSON сетки: %v", err)
	}

	// Загружаем студентов
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
		return fmt.Errorf("ошибка чтения файла студентов: %v", err)
	}
	
	var students studentsRoot
	if err := json.Unmarshal(studentsData, &students); err != nil {
		return fmt.Errorf("ошибка парсинга JSON студентов: %v", err)
	}

	// Создаём маппинг:
	//  - группа -> список студентов
	//  - группа -> отделение (из students.json, чтобы не гадать по префиксам)
	groupStudents := make(map[string][]string)
	groupDepartments := make(map[string]string)
	for _, dept := range students.Departments {
		for _, grp := range dept.Groups {
			studentsList := []string{}
			for _, st := range grp.Students {
				if st.FullName != "" {
					studentsList = append(studentsList, st.FullName)
				}
			}
			key := strings.ToLower(grp.Group)
			groupStudents[key] = studentsList
			if dept.Department != "" {
				groupDepartments[key] = dept.Department
			}
		}
	}

	// Маппинг дней недели на смещение от понедельника
	dayOffset := map[string]int{
		"Понедельник": 0,
		"Вторник":     1,
		"Среда":       2,
		"Четверг":     3,
		"Пятница":     4,
		"Суббота":     5,
		"Воскресенье": 6,
	}

	// Определяем дату начала недели
	var startDate time.Time
	if grid.WeekStartDate != "" {
		parsed, err := time.Parse("2006-01-02", grid.WeekStartDate)
		if err == nil {
			startDate = parsed
		}
	}
	if startDate.IsZero() && weekStartDate != "" {
		parsed, err := time.Parse("2006-01-02", weekStartDate)
		if err == nil {
			startDate = parsed
		}
	}
	if startDate.IsZero() {
		now := time.Now()
		weekday := int(now.Weekday())
		if weekday == 0 {
			weekday = 7
		}
		startDate = now.AddDate(0, 0, -weekday+1)
	}

	// Преобразуем сетку в формат lessons
	groupsMap := make(map[string]*GroupLessons)
	
	for _, record := range grid.Records {
		// Вычисляем дату занятия
		offset, ok := dayOffset[record.DayOfWeek]
		if !ok {
			continue
		}
		lessonDate := startDate.AddDate(0, 0, offset)
		dateStr := lessonDate.Format("02.01.2006 0:00:00")

		// Получаем или создаём группу
		groupKey := strings.ToLower(record.Group)
		if _, exists := groupsMap[groupKey]; !exists {
			deptName := groupDepartments[groupKey]
			if deptName == "" {
				// Фоллбэк на heuristics, если группу не нашли в students.json
				deptName = departmentForGroup(record.Group)
			}
			groupsMap[groupKey] = &GroupLessons{
				Group:         record.Group,
				Department:    deptName,
				Students:      []StudentLessons{},
				TotalStudents: 0,
			}
		}

		// Получаем список студентов для группы
		studentsList, ok := groupStudents[groupKey]
		if !ok {
			// Если группы нет в students.json, создаём пустую запись
			studentsList = []string{}
		}

		// Создаём записи для каждого студента группы
		for _, studentName := range studentsList {
			// Ищем существующего студента
			var studentLessons *StudentLessons
			groupObj := groupsMap[groupKey]
			for j := range groupObj.Students {
				if groupObj.Students[j].StudentName == studentName {
					studentLessons = &groupObj.Students[j]
					break
				}
			}

			// Если студента нет - создаём нового
			if studentLessons == nil {
				groupObj.Students = append(groupObj.Students, StudentLessons{
					StudentName:   studentName,
					NumberInGroup: 0,
					Records:       []LessonRecord{},
					TotalCount:    0,
				})
				studentLessons = &groupObj.Students[len(groupObj.Students)-1]
				groupObj.TotalStudents++
			}

			// Добавляем запись о занятии
			// Если преподаватель не был извлечен при парсинге сетки, пробуем извлечь из дисциплины
			teacher := record.Teacher
			discipline := record.Discipline
			if teacher == "" && discipline != "" {
				// Пробуем извлечь преподавателя из дисциплины
				parsedDisc, parsedTeacher, _ := parseLessonCell(discipline)
				if parsedTeacher != "" {
					teacher = parsedTeacher
					discipline = parsedDisc
				}
			}
			
			studentLessons.Records = append(studentLessons.Records, LessonRecord{
				Date:         dateStr,
				LessonNumber: record.LessonNumber,
				Discipline:   discipline,
				Teacher:      teacher,
				Attendance:   false, // По умолчанию отсутствует, будет заполнено из attendance.json
			})
			studentLessons.TotalCount++
		}
	}

	// Преобразуем map в slice
	groups := make([]GroupLessons, 0, len(groupsMap))
	totalStudents := 0
	for _, g := range groupsMap {
		g.TotalStudents = len(g.Students)
		totalStudents += g.TotalStudents
		groups = append(groups, *g)
	}

	// Формируем итоговую структуру
	output := LessonsOutput{
		Period:        grid.Period,
		Groups:        groups,
		TotalGroups:   len(groups),
		TotalStudents: totalStudents,
	}

	outputPath, err := filepath.Abs(outputFile)
	if err != nil {
		return fmt.Errorf("ошибка получения пути: %v", err)
	}

	jsonData, err := json.MarshalIndent(output, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %v", err)
	}

	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %v", err)
	}

	fmt.Printf(" Преобразование сетки расписания в формат lessons завершено.\n")
	fmt.Printf("   Период: %s\n", output.Period)
	fmt.Printf("   Групп: %d\n", output.TotalGroups)
	fmt.Printf("   Студентов: %d\n", output.TotalStudents)
	fmt.Printf("   Файл сохранён: %s\n", outputPath)

	return nil
}

// parseLessonNumber извлекает номер пары из строки
func parseLessonNumber(s string) int {
	s = strings.TrimSpace(s)
	// Пробуем извлечь число
	var num int
	_, err := fmt.Sscanf(s, "%d", &num)
	if err == nil && num > 0 && num <= 10 {
		return num
	}
	return 0
}

// parseLessonCell парсит содержимое ячейки с информацией о паре
// Формат: "1. [Дисциплина] [Преподаватель] [Аудитория]"
// или многострочный формат с переносами строк
// или однострочный формат: "Дисциплина Фамилия И.О. Аудитория"
func parseLessonCell(cellValue string) (discipline, teacher, location string) {
	cellValue = strings.TrimSpace(cellValue)
	if cellValue == "" {
		return "", "", ""
	}

	// Убираем номер в начале (если есть "1. ", "2. " и т.д.)
	re := regexp.MustCompile(`^\d+\.\s*`)
	cellValue = re.ReplaceAllString(cellValue, "")

	// Разбиваем по переносам строк (могут быть разные символы переноса)
	cellValue = strings.ReplaceAll(cellValue, "\r\n", "\n")
	cellValue = strings.ReplaceAll(cellValue, "\r", "\n")
	lines := strings.Split(cellValue, "\n")
	
	// Фильтруем пустые строки
	nonEmptyLines := []string{}
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed != "" {
			nonEmptyLines = append(nonEmptyLines, trimmed)
		}
	}

	if len(nonEmptyLines) == 0 {
		return "", "", ""
	}

	// Если многострочный формат
	if len(nonEmptyLines) > 1 {
		// Первая строка обычно дисциплина
		discipline = nonEmptyLines[0]
		
		// Ищем преподавателя (обычно строка с инициалами типа "Иванов И.И." или "Иванов И. И.")
		teacherPattern := regexp.MustCompile(`^[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?$`)
		locationPattern := regexp.MustCompile(`(Аудитория|корпус|Спортивный зал)`)
		
		for i := 1; i < len(nonEmptyLines); i++ {
			line := nonEmptyLines[i]
			
			// Проверяем, является ли строка преподавателем
			if teacherPattern.MatchString(line) && teacher == "" {
				teacher = line
				continue
			}
			
			// Проверяем, является ли строка аудиторией
			if locationPattern.MatchString(line) || strings.Contains(line, "Аудитория") || strings.Contains(line, "корпус") {
				if location == "" {
					location = line
				} else {
					location += ", " + line
				}
				continue
			}
			
			// Если это не преподаватель и не аудитория, возможно это продолжение дисциплины
			if teacher == "" && location == "" {
				discipline += " " + line
			}
		}
		return strings.TrimSpace(discipline), strings.TrimSpace(teacher), strings.TrimSpace(location)
	}

	// Однострочный формат - парсим как одну строку
	// Паттерн: "Дисциплина Фамилия И.О. Аудитория" или "Дисциплина Фамилия И.О. -"
	// Преподаватель обычно имеет формат: "Фамилия И.О." где И.О. - инициалы с точками
	
	// Паттерн для поиска преподавателя: Фамилия (заглавная буква + строчные) + пробел + И.О. (заглавная.заглавная. или заглавная. заглавная.)
	// Примеры: "Бурак А.Н.", "Трушина И.Ю.", "Пономаренко Н.В."
	teacherPattern := regexp.MustCompile(`\b([А-ЯЁ][а-яё]+)\s+([А-ЯЁ]\.\s*[А-ЯЁ]\.?)\b`)
	
	// Ищем все совпадения паттерна преподавателя
	matches := teacherPattern.FindAllStringSubmatch(cellValue, -1)
	if len(matches) > 0 {
		// Берем последнее совпадение (обычно преподаватель в конце дисциплины, перед аудиторией)
		match := matches[len(matches)-1]
		teacherFull := match[0] // Полное совпадение: "Фамилия И.О."
		teacher = teacherFull
		
		// Находим позицию преподавателя в строке
		teacherIdx := strings.LastIndex(cellValue, teacherFull)
		if teacherIdx >= 0 {
			// Дисциплина - всё до преподавателя
			discipline = strings.TrimSpace(cellValue[:teacherIdx])
			
			// Остаток после преподавателя - аудитория
			remainder := strings.TrimSpace(cellValue[teacherIdx+len(teacherFull):])
			
			// Обрабатываем остаток
			if remainder != "" && remainder != "-" {
				// Убираем лишние пробелы и разделяем по словам
				remainderParts := strings.Fields(remainder)
				// Проверяем, начинается ли с ключевых слов аудитории
				if strings.Contains(remainder, "Аудитория") || strings.Contains(remainder, "корпус") || 
				   strings.Contains(remainder, "Спортивный зал") || strings.Contains(remainder, "Актовый зал") {
					location = remainder
				} else if len(remainderParts) > 0 {
					// Если остаток не похож на аудиторию, возможно это часть дисциплины
					// Но обычно это аудитория
					location = remainder
				}
			}
		}
	} else {
		// Если не нашли паттерн преподавателя, пробуем найти по другим признакам
		// Ищем слова, которые могут быть фамилиями (заглавная буква + строчные, длина > 3)
		words := strings.Fields(cellValue)
		for i := len(words) - 1; i >= 0; i-- {
			word := words[i]
			// Проверяем, похоже ли слово на инициалы (И.О. или И. О.)
			if matched, _ := regexp.MatchString(`^[А-ЯЁ]\.\s*[А-ЯЁ]\.?$`, word); matched {
				// Если предыдущее слово - фамилия
				if i > 0 {
					prevWord := words[i-1]
					if matched, _ := regexp.MatchString(`^[А-ЯЁ][а-яё]{2,}$`, prevWord); matched {
						teacher = prevWord + " " + word
						discipline = strings.Join(words[:i-1], " ")
						if i+1 < len(words) {
							location = strings.Join(words[i+1:], " ")
						}
						break
					}
				}
			}
		}
		
		// Если всё ещё не нашли, вся строка - дисциплина
		if discipline == "" && teacher == "" {
			discipline = cellValue
		}
	}

	return strings.TrimSpace(discipline), strings.TrimSpace(teacher), strings.TrimSpace(location)
}
