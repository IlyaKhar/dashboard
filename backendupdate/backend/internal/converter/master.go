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
		xlsInfo, xlsErr := os.Stat(inputFileXLS)
		xlsxInfo, xlsxErr := os.Stat(inputFileXLSX)

		needConvert := false
		if os.IsNotExist(xlsxErr) {
			// XLSX не существует, нужно конвертировать
			needConvert = true
		} else if xlsErr == nil && xlsxErr == nil && xlsInfo.ModTime().After(xlsxInfo.ModTime()) {
			// XLS новее XLSX — пересобираем
			needConvert = true
		}

		if needConvert {
			log.Printf("[ConvertMaster] Конвертация XLS → XLSX: %s → %s", inputFileXLS, inputFileXLSX)
			if err := convertXLSToXLSX(inputFileXLS, inputFileXLSX, pythonScriptPath); err != nil {
				// Если конвертация не удалась, пробуем использовать XLS напрямую (не поддерживается excelize)
				result.Warnings = append(result.Warnings, fmt.Sprintf("XLS→XLSX не удалась: %v. Требуется ручная конвертация.", err))
				return nil, fmt.Errorf("не удалось конвертировать XLS в XLSX: %w. Убедитесь, что Python скрипт доступен по пути: %s", err, pythonScriptPath)
			}
			log.Printf("[ConvertMaster] Конвертация успешна: %s", inputFileXLSX)
		} else {
			log.Printf("[ConvertMaster] Используем существующий XLSX (актуален): %s", inputFileXLSX)
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
	// Если в файле нет явной сводной ведомости, но есть детальная посещаемость,
	// строим сводную ведомость агрегированно из attendance (по студентам).
	if len(extracted.VedomostData) == 0 && len(extracted.AttendanceRecords) > 0 {
		log.Printf("[ConvertMaster] В файле не найдена сводная ведомость, строим vedomost.json из attendance...")
		extracted.VedomostData = aggregateAttendanceToVedomost(extracted.AttendanceRecords)
	}

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

			// Обновляем накопительную историю посещаемости (attendance_history.json):
			// добавляем текущий день к уже сохранённым датам или обновляем записи за ту же дату.
			historyPath := filepath.Join(outputDir, "attendance_history.json")
			if err := UpdateAttendanceHistory(attendancePath, historyPath); err != nil {
				log.Printf("[ConvertMaster] Предупреждение: не удалось обновить историю посещаемости: %v", err)
			} else {
				log.Printf("[ConvertMaster] ✓ История посещаемости обновлена: %s", historyPath)
			}
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
	Students          []studentContingentItem
	AttendanceRecords []attendanceRecordItem
	VedomostData      []vedomostItem
	VedomostPeriod    string // период из шапки сводной ведомости (Период: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ)
}

type studentContingentItem struct {
	Department    string
	Group         string
	NumberInGroup int
	FullName      string
	Status        string
}

type attendanceRecordItem struct {
	Department   string
	Group        string
	Student      string
	Date         string // YYYY-MM-DD
	Missed       int
	LessonNumber int // Номер пары (1-6), 0 если определить не удалось
	Discipline   string
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

// aggregateAttendanceToVedomost агрегирует детальную посещаемость по студентам
// в структуру сводной ведомости: накапливаем суммарное количество пропусков.
// Информации о "уважительной/неуважительной" у нас нет, поэтому все часы
// считаем как MissedBad, а MissedTotal = сумма всех пропусков.
func aggregateAttendanceToVedomost(records []attendanceRecordItem) []vedomostItem {
	if len(records) == 0 {
		return nil
	}

	m := make(map[string]*vedomostItem)
	for _, r := range records {
		if r.Missed <= 0 {
			continue
		}
		key := r.Department + "|" + r.Group + "|" + r.Student
		item, ok := m[key]
		if !ok {
			item = &vedomostItem{
				Department:    r.Department,
				Specialty:     "",
				Group:         r.Group,
				Student:       r.Student,
				MissedTotal:   0,
				MissedBad:     0,
				MissedExcused: 0,
			}
			m[key] = item
		}
		item.MissedTotal += r.Missed
		item.MissedBad += r.Missed
	}

	out := make([]vedomostItem, 0, len(m))
	for _, v := range m {
		out = append(out, *v)
	}
	return out
}

// extractAllData умный парсинг файла: определяет тип данных и извлекает всё
func extractAllData(rows [][]string) ExtractedData {
	var data ExtractedData

	var currentDepartment string
	var currentSpecialty string
	var currentGroup string
	var currentLessonNumber int
	var currentAttendanceDiscipline string
	lessonColumnByNumber := make(map[int]int) // lessonNumber -> columnIndex in row

	// Дата по умолчанию для записей посещаемости — берём из периода
	// "Период: 04.03.2026 - 04.03.2026" → 2026-03-04.
	attendanceDefaultDate := ""

	// Проходим по всем строкам и классифицируем их
	for _, row := range rows {
		if len(row) == 0 {
			continue
		}

		firstCell := strings.TrimSpace(row[0])
		if firstCell == "" {
			continue
		}

		// В ведомости номер пары может быть указан не в строке студента,
		// а в строке-заголовке секции (например: "1-я пара", "2 пара").
		// Держим текущую пару и применяем её к строкам с посещаемостью.
		if ln := extractLessonNumberFromRow(row); ln > 0 {
			currentLessonNumber = ln
		}

		// Строим соответствие "номер пары -> индекс колонки" только по строке-шапке.
		// Это важно, потому что слово "пара" может встречаться и в других строках
		// (например, заголовки/вставки), и мы не хотим перетирать маппинг случайными значениями.
		if isLessonHeaderRow(row) {
			for k := range lessonColumnByNumber {
				delete(lessonColumnByNumber, k)
			}
			updateLessonColumnsFromRow(row, lessonColumnByNumber)
		}

		// Извлекаем период из строки "Параметры:"
		if firstCell == "Параметры:" {
			if p := extractPeriodFromRow(row); p != "" {
				data.VedomostPeriod = p
				// Пытаемся вытащить левую дату периода как дату для детальной посещаемости
				if attendanceDefaultDate == "" {
					parts := strings.Split(p, "-")
					if len(parts) > 0 {
						from := strings.TrimSpace(parts[0])
						if d := parseDateValue(from); d != "" {
							attendanceDefaultDate = d
						}
					}
				}
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
			currentLessonNumber = 0
			currentAttendanceDiscipline = ""
			continue
		}

		if isSpecialty(firstCell) {
			currentSpecialty = firstCell
			currentGroup = ""
			currentLessonNumber = 0
			currentAttendanceDiscipline = ""
			continue
		}

		if isGroup(firstCell) {
			currentGroup = strings.ToLower(firstCell)
			currentLessonNumber = 0
			currentAttendanceDiscipline = ""
			continue
		}

		// В текущем формате ведомости значения вроде "2.0/4.0/6.0" в первой колонке
		// не считаем надёжным номером пары: они дают ложное распределение пропусков по урокам.
		// Номер пары берём только из явно размеченных строк (через parseAttendanceRow/extractLessonNumberFromRow).

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

		// Вариант 2: Строка детальной посещаемости (есть пропущенные часы;
		// дату берём либо из строки, либо из периода, если в самой строке её нет).
		if currentDepartment != "" && currentGroup != "" {
			// Запоминаем "текущую дисциплину" для последующих строк студентов.
			// Ведомость часто идёт блоками: строка с названием дисциплины, затем ФИО студентов.
			if !looksLikeFullName(firstCell) && !isHeaderOrTotal(firstCell) {
				if missedTotal, _, _ := parseVedomostRow(row); missedTotal > 0 {
					currentAttendanceDiscipline = firstCell
				}
			}

			// Дата: из первой ячейки, если она похожа на дату, иначе из периода.
			date := parseDateValue(firstCell)
			if date == "" {
				date = attendanceDefaultDate
			}

			if date != "" {
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
					// Если мы смогли найти колонки пар в шапке — парсим пропуски по каждой колонке,
					// чтобы значения для пары 1 и пары 2 не были одинаковыми.
					if len(lessonColumnByNumber) > 0 {
						addedAny := false
						for _, colIdx := range lessonColumnByNumber {
							if colIdx < 0 || colIdx >= len(row) {
								continue
							}
							missed := parseIntCell(row[colIdx])
							if missed <= 0 {
								continue
							}

							addedAny = true
							data.AttendanceRecords = append(data.AttendanceRecords, attendanceRecordItem{
								Department:   currentDepartment,
								Group:        currentGroup,
								Student:      studentName,
								Date:         date,
								Missed:       missed,
							// Номер пары в текущей ведомости считаем ненадёжным:
							// не распределяем пропуски по парам, оставляем 0.
							LessonNumber: 0,
								Discipline:   currentAttendanceDiscipline,
							})
						}

						if addedAny {
							continue
						}
					}

					// Fallback: старое поведение (одно число на строку) — чтобы не потерять данные
					// в форматах ведомости, где шапка с колонками пар не распознаётся.
					if _, missed, lessonNumber := parseAttendanceRow(row, attendanceDefaultDate); missed > 0 {
						_ = lessonNumber
						_ = currentLessonNumber
						data.AttendanceRecords = append(data.AttendanceRecords, attendanceRecordItem{
							Department:   currentDepartment,
							Group:        currentGroup,
							Student:      studentName,
							Date:         date,
							Missed:       missed,
							// Номер пары в текущей ведомости считаем ненадёжным:
							// не распределяем пропуски по парам, оставляем 0.
							LessonNumber: 0,
							Discipline:   currentAttendanceDiscipline,
						})
					}
				}
			}
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
						Specialty:     currentSpecialty,
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

// parseVedomostRow пытается найти пропуски из сводной ведомости
func parseVedomostRow(row []string) (total, bad, excused int) {
	// 1. Пытаемся распарсить "старый" шаблон: A (ФИО), D (неуваж), E (уваж), H (всего).
	if len(row) > 7 {
		total = parseIntCell(row[7])
	}
	if len(row) > 4 {
		excused = parseIntCell(row[4])
	}
	if len(row) > 3 {
		bad = parseIntCell(row[3])
	}
	if total > 0 || excused > 0 || bad > 0 {
		if total == 0 && excused > 0 {
			total = excused
		}
		return total, bad, excused
	}

	// 2. Универсальный режим для "новых" ведомостей: берём последние числовые ячейки строки.
	// Ищем все числа в строке, кроме первой колонки (там обычно текст: ФИО/группа).
	nums := make([]int, 0, 4)
	for i := 1; i < len(row); i++ {
		if v := parseIntCell(row[i]); v > 0 {
			nums = append(nums, v)
		}
	}
	if len(nums) == 0 {
		return 0, 0, 0
	}

	// Считаем, что последнее число в строке — "всего пропущено часов".
	total = nums[len(nums)-1]

	// Предпоследнее — "по уважительной", пред‑предпоследнее — "по неуважительной" (если есть).
	if len(nums) >= 2 {
		excused = nums[len(nums)-2]
	}
	if len(nums) >= 3 {
		bad = nums[len(nums)-3]
	}

	// Если распределение странное, стараемся не ломать total.
	if bad+excused > total && total > 0 {
		bad = 0
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

// extractLessonNumberFromRow пробует извлечь номер пары из любой ячейки строки.
// Возвращает 0, если в строке явно не указано "пара" (чтобы не ловить случайные цифры).
func extractLessonNumberFromRow(row []string) int {
	for _, cell := range row {
		text := strings.ToLower(strings.TrimSpace(cell))
		if text == "" {
			continue
		}

		// Формат ведомости: в строках блока "Номер пары" значение может быть просто "1", "2", ... "6".
		if text == "1" || text == "2" || text == "3" || text == "4" || text == "5" || text == "6" {
			n, _ := strconv.Atoi(text)
			return n
		}

		// Поддержка текстовых форм: "1-я пара", "пара 2", и т.д.
		if strings.Contains(text, "пара") {
			for _, r := range text {
				if r >= '1' && r <= '6' {
					return int(r - '0')
				}
			}
		}
	}
	return 0
}

func updateLessonColumnsFromRow(row []string, out map[int]int) {
	for colIdx, cell := range row {
		ln := extractLessonNumberFromHeaderCell(cell)
		if ln <= 0 {
			continue
		}
		out[ln] = colIdx
	}
}

// extractLessonNumberFromHeaderCell пытается распознать запись вида "1-я пара" или "2 пара".
// Возвращает 0, если в ячейке не похоже на заголовок пары.
func extractLessonNumberFromHeaderCell(cell string) int {
	text := strings.ToLower(strings.TrimSpace(cell))
	if text == "" {
		return 0
	}

	// Номер пары считаем только из явно размеченных ячеек со словом "пара".
	// Иначе в сводной ведомости короткие числовые значения ("2.0", "4.0", "6.0")
	// начинают ошибочно восприниматься как номер пары.
	if strings.Contains(text, "пара") {
		for _, r := range text {
			if r >= '1' && r <= '6' {
				return int(r - '0')
			}
		}
	}
	return 0
}

func isLessonHeaderRow(row []string) bool {
	// Заголовок пары обычно содержит сразу несколько значений вида "1-я/2-я пара".
	// Мы ищем не менее двух разных lessonNumber в одной строке.
	seen := make(map[int]struct{})
	for _, cell := range row {
		ln := extractLessonNumberFromHeaderCell(cell)
		if ln <= 0 {
			continue
		}
		seen[ln] = struct{}{}
		if len(seen) >= 2 {
			return true
		}
	}
	return false
}

func looksLikeFullName(text string) bool {
	parts := strings.Fields(strings.TrimSpace(text))
	if len(parts) != 3 {
		return false
	}
	for _, p := range parts {
		if len([]rune(p)) < 2 {
			return false
		}
		for _, r := range p {
			if (r >= '0' && r <= '9') || r == '/' {
				return false
			}
		}
	}
	return true
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
			if rec.Date == item.Date && rec.LessonNumber == item.LessonNumber && rec.Discipline == item.Discipline {
				dateExists = true
				break
			}
		}
		if !dateExists {
			studentObj.Attendance = append(studentObj.Attendance, AttendanceRecord{
				Date:         item.Date,
				Missed:       item.Missed,
				LessonNumber: item.LessonNumber,
				Discipline:   item.Discipline,
			})
		} else {
			// Если запись за эту дату и эту пару уже есть, суммируем пропуски.
			for i := range studentObj.Attendance {
				if studentObj.Attendance[i].Date == item.Date &&
					studentObj.Attendance[i].LessonNumber == item.LessonNumber &&
					studentObj.Attendance[i].Discipline == item.Discipline {
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
