package converter

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// MasterConversionResult результат мастер-конвертации из одного файла ведомости
type MasterConversionResult struct {
	StudentsOutput   string // путь к students.json
	AttendanceOutput string // путь к attendance.json
	VedomostOutput   string // путь к vedomost.json
	Errors           []string
	Warnings         []string
}

// ConvertMaster универсальный мастер-конвертер: один файл ведомость.xls → все JSON
// Анализирует структуру файла и извлекает:
//   - Контингент студентов (students.json)
//   - Детальную посещаемость по датам (attendance.json)
//   - Сводную ведомость пропусков (vedomost.json)
func ConvertMaster(
	inputFileXLS string,
	outputDir string,
	pythonScriptPath string,
) (*MasterConversionResult, error) {
	result := &MasterConversionResult{
		Errors:   []string{},
		Warnings: []string{},
	}

	// Шаг 1: Конвертируем XLS → XLSX (если нужно)
	inputFileXLSX := inputFileXLS
	if strings.HasSuffix(strings.ToLower(inputFileXLS), ".xls") {
		// Проверяем, существует ли исходный XLS файл
		if _, err := os.Stat(inputFileXLS); os.IsNotExist(err) {
			return nil, fmt.Errorf("файл не найден: %s", inputFileXLS)
		}
		
		// Пробуем найти уже существующий XLSX
		inputFileXLSX = strings.TrimSuffix(inputFileXLS, ".xls") + ".xlsx"
		if _, err := os.Stat(inputFileXLSX); os.IsNotExist(err) {
			// XLSX не существует, нужно конвертировать
			log.Printf("[ConvertMaster] Конвертация XLS → XLSX: %s → %s", inputFileXLS, inputFileXLSX)
			if err := convertXLSToXLSX(inputFileXLS, inputFileXLSX, pythonScriptPath); err != nil {
				// Если конвертация не удалась, пробуем использовать XLS напрямую (не поддерживается excelize)
				result.Warnings = append(result.Warnings, fmt.Sprintf("XLS→XLSX не удалась: %v. Требуется ручная конвертация.", err))
				return nil, fmt.Errorf("не удалось конвертировать XLS в XLSX: %w. Убедитесь, что Python скрипт доступен по пути: %s", err, pythonScriptPath)
			}
			log.Printf("[ConvertMaster] Конвертация успешна: %s", inputFileXLSX)
		} else {
			log.Printf("[ConvertMaster] XLSX файл уже существует: %s", inputFileXLSX)
		}
	}

	// Шаг 2: Открываем XLSX
	log.Printf("[ConvertMaster] Открываем файл: %s", inputFileXLSX)
	f, err := excelize.OpenFile(inputFileXLSX)
	if err != nil {
		return nil, fmt.Errorf("ошибка открытия файла %s: %w", inputFileXLSX, err)
	}
	defer f.Close()

	// Шаг 3: Анализируем структуру файла
	// Пробуем определить тип данных по первому листу
	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return nil, fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("ошибка чтения строк: %w", err)
	}

	// Шаг 4: Извлекаем все данные из файла
	extracted := extractAllData(rows)

	// Шаг 5: Генерируем JSON файлы
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return nil, fmt.Errorf("ошибка создания директории %s: %w", outputDir, err)
	}

	// 5.1. students.json (контингент)
	if len(extracted.Students) > 0 {
		studentsPath := filepath.Join(outputDir, "students.json")
		if err := writeStudentsJSON(extracted.Students, studentsPath); err != nil {
			result.Errors = append(result.Errors, fmt.Sprintf("students.json: %v", err))
		} else {
			result.StudentsOutput = studentsPath
		}
	} else {
		result.Warnings = append(result.Warnings, "Не найдены данные контингента студентов")
	}

	// 5.2. attendance.json (детальная посещаемость по датам)
	if len(extracted.AttendanceRecords) > 0 {
		attendancePath := filepath.Join(outputDir, "attendance.json")
		if err := writeAttendanceJSON(extracted.AttendanceRecords, attendancePath); err != nil {
			result.Errors = append(result.Errors, fmt.Sprintf("attendance.json: %v", err))
		} else {
			result.AttendanceOutput = attendancePath
		}
	} else {
		result.Warnings = append(result.Warnings, "Не найдены записи детальной посещаемости")
	}

	// 5.3. vedomost.json (сводная ведомость)
	if len(extracted.VedomostData) > 0 {
		vedomostPath := filepath.Join(outputDir, "vedomost.json")
		if err := writeVedomostJSON(extracted.VedomostData, extracted.VedomostPeriod, vedomostPath); err != nil {
			result.Errors = append(result.Errors, fmt.Sprintf("vedomost.json: %v", err))
		} else {
			result.VedomostOutput = vedomostPath
		}
	} else {
		result.Warnings = append(result.Warnings, "Не найдены данные сводной ведомости")
	}

	return result, nil
}

// ExtractedData все данные, извлечённые из файла
type ExtractedData struct {
	Students         []studentContingentItem
	AttendanceRecords []attendanceRecordItem
	VedomostData     []vedomostItem
	VedomostPeriod   string // период из шапки сводной ведомости (Период: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ)
}

type studentContingentItem struct {
	Department    string
	Group         string
	NumberInGroup int
	FullName      string
	Status        string
}

type attendanceRecordItem struct {
	Department string
	Group      string
	Student    string
	Date       string // YYYY-MM-DD
	Missed     int
}

type vedomostItem struct {
	Department    string
	Specialty     string
	Group         string
	Student       string
	MissedTotal   int
	MissedBad     int
	MissedExcused int
}

// extractAllData умный парсинг файла: определяет тип данных и извлекает всё
func extractAllData(rows [][]string) ExtractedData {
	var data ExtractedData

	var currentDepartment string
	var currentSpecialty string
	var currentGroup string

	// Проходим по всем строкам и классифицируем их
	for _, row := range rows {
		if len(row) == 0 {
			continue
		}

		firstCell := strings.TrimSpace(row[0])
		if firstCell == "" {
			continue
		}

		// Извлекаем период из строки "Параметры:"
		if firstCell == "Параметры:" {
			if p := extractPeriodFromRow(row); p != "" {
				data.VedomostPeriod = p
			}
			continue
		}

		// Пропускаем заголовки
		if isHeaderOrTotal(firstCell) {
			continue
		}

		// Определяем тип строки
		if isDepartment(firstCell) {
			currentDepartment = firstCell
			currentSpecialty = ""
			currentGroup = ""
			continue
		}

		if isSpecialty(firstCell) {
			currentSpecialty = firstCell
			currentGroup = ""
			continue
		}

		if isGroup(firstCell) {
			currentGroup = strings.ToLower(firstCell)
			continue
		}

		// Пытаемся определить, что это за строка
		// Вариант 1: Строка контингента (есть № в группе и ФИО)
		if numInGroup, fullName := parseStudentRow(row); numInGroup > 0 && fullName != "" {
			if currentDepartment != "" && currentGroup != "" {
				data.Students = append(data.Students, studentContingentItem{
					Department:    currentDepartment,
					Group:         currentGroup,
					NumberInGroup: numInGroup,
					FullName:      fullName,
					Status:        "Студент",
				})
			}
			continue
		}

		// Вариант 2: Строка детальной посещаемости (есть дата и пропущенные часы)
		if date, missed := parseAttendanceRow(row); date != "" && missed > 0 {
			if currentDepartment != "" && currentGroup != "" {
				// Ищем ФИО студента в строке
				studentName := findStudentNameInRow(row)
				if studentName == "" {
					// Если не нашли в строке, берём последнего студента из контингента
					if len(data.Students) > 0 {
						last := data.Students[len(data.Students)-1]
						if last.Department == currentDepartment && last.Group == currentGroup {
							studentName = last.FullName
						}
					}
				}
				if studentName != "" {
					data.AttendanceRecords = append(data.AttendanceRecords, attendanceRecordItem{
						Department: currentDepartment,
						Group:      currentGroup,
						Student:    studentName,
						Date:       date,
						Missed:     missed,
					})
				}
			}
			continue
		}

		// Вариант 3: Строка сводной ведомости (есть пропуски по уважительным/неуважительным)
		if missedTotal, missedBad, missedExcused := parseVedomostRow(row); missedTotal > 0 {
			if currentDepartment != "" && currentGroup != "" {
				studentName := findStudentNameInRow(row)
				if studentName == "" {
					// Если не нашли, используем текст из первой колонки как имя студента
					words := strings.Fields(firstCell)
					if len(words) >= 2 {
						studentName = firstCell
					}
				}
				if studentName != "" {
					data.VedomostData = append(data.VedomostData, vedomostItem{
						Department:    currentDepartment,
						Specialty:      currentSpecialty,
						Group:         currentGroup,
						Student:       studentName,
						MissedTotal:   missedTotal,
						MissedBad:     missedBad,
						MissedExcused: missedExcused,
					})
				}
			}
		}
	}

	return data
}

// parseStudentRow пытается найти № в группе и ФИО студента
func parseStudentRow(row []string) (numberInGroup int, fullName string) {
	// Ищем номер в группе (колонки B, C, D)
	for i := 1; i <= 3 && i < len(row); i++ {
		if num, err := strconv.Atoi(strings.TrimSpace(row[i])); err == nil && num > 0 {
			numberInGroup = num
			break
		}
	}

	// Ищем ФИО (колонка E или дальше, минимум 2 слова)
	for i := 4; i < len(row) && i < 10; i++ {
		val := strings.TrimSpace(row[i])
		words := strings.Fields(val)
		if len(words) >= 2 && len(words) <= 4 {
			fullName = val
			break
		}
	}

	return numberInGroup, fullName
}

// parseAttendanceRow пытается найти дату и пропущенные часы
func parseAttendanceRow(row []string) (date string, missed int) {
	// Ищем дату в первой колонке или в колонке F
	for _, idx := range []int{0, 5} {
		if idx < len(row) {
			val := strings.TrimSpace(row[idx])
			if parsed := parseDateValue(val); parsed != "" {
				date = parsed
				break
			}
		}
	}

	// Ищем пропущенные часы (колонка F или последняя числовая колонка)
	for i := len(row) - 1; i >= 0 && i >= 5; i-- {
		if num, err := strconv.ParseFloat(strings.TrimSpace(row[i]), 64); err == nil && num > 0 {
			missed = int(num)
			break
		}
	}

	return date, missed
}

// parseVedomostRow пытается найти пропуски из сводной ведомости
func parseVedomostRow(row []string) (total, bad, excused int) {
	// Структура: A (текст), B-C (пусто), D (неуваж), E (уваж), F-G (пусто), H (всего)
	if len(row) > 7 {
		total = parseIntCell(row[7])
	}
	if len(row) > 4 {
		excused = parseIntCell(row[4])
	}
	if len(row) > 3 {
		bad = parseIntCell(row[3])
	}

	if total == 0 && excused > 0 {
		total = excused
	}

	return total, bad, excused
}

// findStudentNameInRow ищет ФИО студента в строке
func findStudentNameInRow(row []string) string {
	for i := 0; i < len(row) && i < 10; i++ {
		val := strings.TrimSpace(row[i])
		words := strings.Fields(val)
		if len(words) >= 2 && len(words) <= 4 {
			// Проверяем, что это не группа и не отделение
			if !isGroup(val) && !isDepartment(val) && !isSpecialty(val) {
				return val
			}
		}
	}
	return ""
}

// writeStudentsJSON генерирует students.json из извлечённых данных контингента
func writeStudentsJSON(items []studentContingentItem, outputPath string) error {
	departmentsMap := make(map[string]*DepartmentContingent)

	for _, item := range items {
		dept, deptExists := departmentsMap[item.Department]
		if !deptExists {
			dept = &DepartmentContingent{
				Department:    item.Department,
				Groups:        []GroupContingent{},
				TotalStudents: 0,
			}
			departmentsMap[item.Department] = dept
		}

		var groupObj *GroupContingent
		for i := range dept.Groups {
			if dept.Groups[i].Group == item.Group {
				groupObj = &dept.Groups[i]
				break
			}
		}
		if groupObj == nil {
			dept.Groups = append(dept.Groups, GroupContingent{
				Group:    item.Group,
				Students: []StudentInfo{},
			})
			groupObj = &dept.Groups[len(dept.Groups)-1]
		}

		// Проверяем, нет ли уже такого студента
		var studentExists bool
		for _, s := range groupObj.Students {
			if s.FullName == item.FullName {
				studentExists = true
				break
			}
		}
		if !studentExists {
			groupObj.Students = append(groupObj.Students, StudentInfo{
				SerialNumber:  item.NumberInGroup,
				NumberInGroup: item.NumberInGroup,
				FullName:      item.FullName,
				Status:        item.Status,
			})
		}
	}

	// Подсчитываем итоги
	totalStudents := 0
	departments := make([]DepartmentContingent, 0, len(departmentsMap))
	for _, d := range departmentsMap {
		deptTotal := 0
		for _, g := range d.Groups {
			deptTotal += len(g.Students)
			totalStudents += len(g.Students)
		}
		d.TotalStudents = deptTotal
		departments = append(departments, *d)
	}

	output := StudentsOutput{
		TotalStudents: totalStudents,
		Departments:   departments,
	}

	jsonData, err := json.MarshalIndent(output, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %w", err)
	}

	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %w", err)
	}

	return nil
}

// writeAttendanceJSON генерирует attendance.json из извлечённых записей посещаемости
func writeAttendanceJSON(items []attendanceRecordItem, outputPath string) error {
	departmentsMap := make(map[string]*Department)

	for _, item := range items {
		dept, exists := departmentsMap[item.Department]
		if !exists {
			dept = &Department{
				Department: item.Department,
				Groups:     []Group{},
			}
			departmentsMap[item.Department] = dept
		}

		var groupObj *Group
		for i := range dept.Groups {
			if dept.Groups[i].Group == item.Group {
				groupObj = &dept.Groups[i]
				break
			}
		}
		if groupObj == nil {
			dept.Groups = append(dept.Groups, Group{
				Group:    item.Group,
				Students: []Student{},
			})
			groupObj = &dept.Groups[len(dept.Groups)-1]
		}

		var studentObj *Student
		for i := range groupObj.Students {
			if groupObj.Students[i].Student == item.Student {
				studentObj = &groupObj.Students[i]
				break
			}
		}
		if studentObj == nil {
			groupObj.Students = append(groupObj.Students, Student{
				Student:    item.Student,
				Attendance: []AttendanceRecord{},
			})
			studentObj = &groupObj.Students[len(groupObj.Students)-1]
		}

		// Проверяем, нет ли уже записи за эту дату
		dateExists := false
		for _, rec := range studentObj.Attendance {
			if rec.Date == item.Date {
				dateExists = true
				break
			}
		}
		if !dateExists {
			studentObj.Attendance = append(studentObj.Attendance, AttendanceRecord{
				Date:   item.Date,
				Missed: item.Missed,
			})
		} else {
			// Если дата уже есть, суммируем пропуски
			for i := range studentObj.Attendance {
				if studentObj.Attendance[i].Date == item.Date {
					studentObj.Attendance[i].Missed += item.Missed
					break
				}
			}
		}
	}

	departments := make([]Department, 0, len(departmentsMap))
	for _, d := range departmentsMap {
		departments = append(departments, *d)
	}

	jsonData, err := json.MarshalIndent(departments, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %w", err)
	}

	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %w", err)
	}

	return nil
}

// writeVedomostJSON генерирует vedomost.json из извлечённых данных сводной ведомости (с полем period)
func writeVedomostJSON(items []vedomostItem, period string, outputPath string) error {
	departmentsMap := make(map[string]*DepartmentSummary)

	for _, item := range items {
		dept, exists := departmentsMap[item.Department]
		if !exists {
			dept = &DepartmentSummary{
				Department:  item.Department,
				TotalMissed: 0,
				Specialties: []SpecialtySummary{},
			}
			departmentsMap[item.Department] = dept
		}

		var spec *SpecialtySummary
		for i := range dept.Specialties {
			if dept.Specialties[i].Specialty == item.Specialty {
				spec = &dept.Specialties[i]
				break
			}
		}
		if spec == nil {
			dept.Specialties = append(dept.Specialties, SpecialtySummary{
				Specialty:   item.Specialty,
				TotalMissed: 0,
				Groups:      []GroupSummary{},
			})
			spec = &dept.Specialties[len(dept.Specialties)-1]
		}

		var group *GroupSummary
		for i := range spec.Groups {
			if spec.Groups[i].Group == item.Group {
				group = &spec.Groups[i]
				break
			}
		}
		if group == nil {
			spec.Groups = append(spec.Groups, GroupSummary{
				Group:       item.Group,
				TotalMissed: 0,
				Students:    []StudentSummary{},
			})
			group = &spec.Groups[len(spec.Groups)-1]
		}

		// Проверяем, нет ли уже такого студента
		studentExists := false
		for i := range group.Students {
			if group.Students[i].Student == item.Student {
				group.Students[i].MissedTotal += item.MissedTotal
				group.Students[i].MissedBad += item.MissedBad
				group.Students[i].MissedExcused += item.MissedExcused
				studentExists = true
				break
			}
		}
		if !studentExists {
			group.Students = append(group.Students, StudentSummary{
				Student:       item.Student,
				MissedTotal:   item.MissedTotal,
				MissedBad:     item.MissedBad,
				MissedExcused: item.MissedExcused,
			})
		}

		// Обновляем суммы
		group.TotalMissed += item.MissedTotal
		spec.TotalMissed += item.MissedTotal
		dept.TotalMissed += item.MissedTotal
	}

	departments := make([]DepartmentSummary, 0, len(departmentsMap))
	for _, d := range departmentsMap {
		departments = append(departments, *d)
	}

	root := vedomostOutput{Period: period, Departments: departments}
	jsonData, err := json.MarshalIndent(root, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %w", err)
	}

	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %w", err)
	}

	return nil
}
