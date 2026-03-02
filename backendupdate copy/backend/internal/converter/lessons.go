package converter

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Типы данных для расписания занятий

// LessonRecord - запись о занятии для одного студента
type LessonRecord struct {
	Date         string `json:"date"`                   // Дата и время занятия (формат: "03.02.2026 0:00:00")
	LessonNumber int    `json:"lessonNumber,omitempty"` // Номер пары (1-6)
	Discipline   string `json:"discipline"`             // Название дисциплины
	Teacher      string `json:"teacher"`                // ФИО преподавателя
	Attendance   bool   `json:"attendance"`             // Явка: true = "Да", false = "Нет"
}

// StudentLessons - занятия для одного студента
type StudentLessons struct {
	StudentName   string         `json:"studentName"`   // ФИО студента
	NumberInGroup int            `json:"numberInGroup"` // Номер по списку
	Records       []LessonRecord `json:"records"`       // Список занятий
	TotalCount    int            `json:"totalCount"`    // Общее количество записей
}

// GroupLessons - занятия для группы
type GroupLessons struct {
	Group         string           `json:"group"`         // Название группы (например, "1a1")
	Department    string           `json:"department"`    // Отделение (определяется по префиксу группы)
	Students      []StudentLessons `json:"students"`      // Список студентов с их занятиями
	TotalStudents int              `json:"totalStudents"` // Общее количество студентов в группе
}

// LessonsOutput - итоговая структура для расписания
type LessonsOutput struct {
	Period        string         `json:"period"`        // Период (например, "02.02.2026 - 03.02.2026")
	Groups        []GroupLessons `json:"groups"`        // Список групп
	TotalGroups   int            `json:"totalGroups"`   // Общее количество групп
	TotalStudents int            `json:"totalStudents"` // Общее количество студентов
}

// ConvertLessons конвертирует файл "Проба.xlsx" (расписание занятий) в JSON
// inputFile - путь к файлу Проба.xlsx
// outputFile - путь к выходному JSON файлу
func ConvertLessons(inputFile, outputFile string) error {
	fmt.Printf("[DEBUG] Открытие файла: %s\n", inputFile)
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		return fmt.Errorf("ошибка открытия файла: %v", err)
	}
	defer f.Close()
	fmt.Printf("[DEBUG] Файл успешно открыт\n")

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("ошибка чтения строк: %v", err)
	}
	fmt.Printf("[DEBUG] Прочитано строк: %d\n", len(rows))

	// Ищем строку с заголовками таблицы
	headerRow := -1
	dataStartRow := -1

	for i, row := range rows {
		if len(row) == 0 {
			continue
		}

		rowText := strings.Join(row, " ")
		// Ищем строку с заголовками: Группа, Студент, Дата, Дисциплина, Преподаватель, Явка
		if strings.Contains(rowText, "Группа") &&
			(strings.Contains(rowText, "Студент") || strings.Contains(rowText, "Студенты")) &&
			strings.Contains(rowText, "Дата") &&
			strings.Contains(rowText, "Дисциплина") {
			headerRow = i
			dataStartRow = i + 1
			break
		}
	}

	if dataStartRow == -1 {
		// Если не нашли заголовки, начинаем с 5-й строки (после параметров)
		dataStartRow = 5
		fmt.Printf("[DEBUG] Заголовки не найдены, начинаем с строки %d\n", dataStartRow)
	} else {
		fmt.Printf("[DEBUG] Заголовки найдены в строке %d, данные начинаются с %d\n", headerRow, dataStartRow)
	}

	// Определяем индексы колонок
	var colGroup, colDiscipline, colTeacher, colAttendance int

	// В этом отчёте структура колонок следующая (по реальным данным из Excel):
	// A (0)  - Группа / Студент / Дата (зависит от типа строки)
	// B (1)  - обычно пустая
	// C (2)  - обычно пустая
	// D (3)  - Дисциплина
	// E (4)  - Преподаватель
	// F (5)  - обычно пустая
	// G (6)  - Студенты.Явка (Да/Нет)
	//
	// Логика:
	// - Если в A код группы (1а1, 2вб3 и т.п.) - это группа
	// - Если в A ФИО (2-4 слова) - это студент
	// - Если в A дата (03.02.2026) - это запись занятия
	// - Дисциплина всегда в D (3)
	// - Преподаватель всегда в E (4)
	// - Явка всегда в G (6)
	colGroup = 0 // A: группа (определяется по isGroupCode)
	// colStudent и colDate тоже в колонке A (0), определяются по содержимому
	colDiscipline = 3 // D: дисциплина
	colTeacher = 4    // E: преподаватель
	colAttendance = 6 // G: явка

	fmt.Printf("[DEBUG] Используем колонки (lessons): Группа/Студент/Дата=0 (колонка A), Дисциплина=%d, Преподаватель=%d, Явка=%d\n",
		colDiscipline, colTeacher, colAttendance)

	// Ищем период в параметрах (первые строки)
	period := ""
	for i := 0; i < dataStartRow && i < len(rows); i++ {
		rowText := strings.Join(rows[i], " ")
		if strings.Contains(rowText, "Период:") {
			// Извлекаем период (например, "02.02.2026 - 03.02.2026")
			parts := strings.Split(rowText, "Период:")
			if len(parts) > 1 {
				period = strings.TrimSpace(parts[1])
			}
			break
		}
	}

	// Структура для хранения данных
	groupsMap := make(map[string]*GroupLessons)

	var currentGroup string
	var currentStudent string
	var currentStudentLessons *StudentLessons

	// Обрабатываем данные
	processedRows := 0
	debugLimit := 100 // Увеличиваем лимит отладки
	for i := dataStartRow; i < len(rows); i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		// Временное логирование первых строк для отладки
		if processedRows < debugLimit {
			fmt.Printf("[DEBUG ROW %d] Колонок: %d, Данные: ", i+1, len(row))
			for j := 0; j < len(row) && j < 7; j++ {
				val := strings.TrimSpace(row[j])
				if len(val) > 20 {
					val = val[:20] + "..."
				}
				fmt.Printf("[%d]'%s' ", j, val)
			}
			fmt.Println()
			processedRows++
		}

		// Пропускаем служебные строки
		firstCell := ""
		if len(row) > 0 {
			firstCell = strings.TrimSpace(row[0])
		}

		if firstCell == "" {
			continue
		}

		// Пропускаем строки с параметрами и заголовками
		if strings.Contains(firstCell, "Параметры") ||
			strings.Contains(firstCell, "Период") ||
			strings.Contains(firstCell, "Тип объекта") ||
			strings.Contains(firstCell, "Имя объекта") ||
			strings.Contains(firstCell, "Имя таблицы") ||
			strings.Contains(firstCell, "Выводить") ||
			strings.Contains(firstCell, "Отделение") ||
			firstCell == "#" ||
			strings.Contains(firstCell, "Количество записей") ||
			firstCell == "Группа" {
			continue
		}

		// Читаем значение из колонки A (может быть группа, студент или дата)
		colAValue := ""
		if colGroup >= 0 && colGroup < len(row) {
			colAValue = strings.TrimSpace(row[colGroup])
		}

		// Определяем тип строки по значению в колонке A
		isGroup := isGroupCode(colAValue)
		isStudent := false
		isDate := false

		if !isGroup && colAValue != "" {
			// Проверяем, является ли это датой
			if strings.Contains(colAValue, ".") && (strings.Contains(colAValue, "2026") || strings.Contains(colAValue, "2025") || strings.Contains(colAValue, "2024")) {
				isDate = true
			} else {
				// Проверяем, является ли это ФИО студента
				words := strings.Fields(colAValue)
				if len(words) >= 2 && len(words) <= 4 {
					if !isGroupCode(colAValue) {
						isStudent = true
					}
				}
			}
		}

		// Обрабатываем группу
		if isGroup {
			currentGroup = colAValue
			currentStudent = ""
			currentStudentLessons = nil

			if _, exists := groupsMap[currentGroup]; !exists {
				groupsMap[currentGroup] = &GroupLessons{
					Group:         currentGroup,
					Students:      []StudentLessons{},
					TotalStudents: 0,
				}
			}
			if processedRows < debugLimit {
				fmt.Printf("[DEBUG] Найдена группа '%s'\n", currentGroup)
			}
			continue
		}

		// Если нет группы - пропускаем
		if currentGroup == "" {
			continue
		}

		// Обрабатываем студента
		if isStudent {
			studentName := colAValue
			currentStudent = studentName
			currentStudentLessons = nil

			// Ищем существующего студента в группе
			groupObj := groupsMap[currentGroup]
			for j := range groupObj.Students {
				if groupObj.Students[j].StudentName == currentStudent {
					currentStudentLessons = &groupObj.Students[j]
					break
				}
			}

			// Если студента нет - создаём нового
			if currentStudentLessons == nil {
				groupObj.Students = append(groupObj.Students, StudentLessons{
					StudentName:   currentStudent,
					NumberInGroup: 0,
					Records:       []LessonRecord{},
					TotalCount:    0,
				})
				currentStudentLessons = &groupObj.Students[len(groupObj.Students)-1]
				groupObj.TotalStudents++
				if processedRows < debugLimit {
					fmt.Printf("[DEBUG] Найден студент '%s' в группе '%s'\n", currentStudent, currentGroup)
				}
			}
			continue
		}

		// Обрабатываем запись занятия (дата в колонке A)
		if !isDate {
			// Если это не дата, не группа и не студент - пропускаем
			continue
		}

		// Если нет студента - пропускаем запись занятия
		if currentStudent == "" || currentStudentLessons == nil {
			if processedRows < debugLimit {
				fmt.Printf("[DEBUG] Пропуск строки %d: нет студента для группы '%s' (дата='%s')\n", i+1, currentGroup, colAValue)
			}
			continue
		}

		// Читаем данные занятия
		date := colAValue // Дата уже в колонке A

		discipline := ""
		if colDiscipline >= 0 && colDiscipline < len(row) {
			discipline = strings.TrimSpace(row[colDiscipline])
		}

		teacher := ""
		if colTeacher >= 0 && colTeacher < len(row) {
			teacher = strings.TrimSpace(row[colTeacher])
		}

		attendanceStr := ""
		if colAttendance >= 0 && colAttendance < len(row) {
			attendanceStr = strings.TrimSpace(row[colAttendance])
		}

		// Пропускаем, если нет даты или дисциплины
		if date == "" || discipline == "" {
			if processedRows < debugLimit {
				fmt.Printf("[DEBUG] Пропуск строки %d: нет даты или дисциплины (date='%s', discipline='%s', student='%s')\n", i+1, date, discipline, currentStudent)
			}
			continue
		}

		// Парсим явку
		attendance := false
		if strings.ToLower(attendanceStr) == "да" || strings.ToLower(attendanceStr) == "yes" {
			attendance = true
		}

		// Добавляем запись о занятии
		currentStudentLessons.Records = append(currentStudentLessons.Records, LessonRecord{
			Date:       date,
			Discipline: discipline,
			Teacher:    teacher,
			Attendance: attendance,
		})
		currentStudentLessons.TotalCount++

		if processedRows < debugLimit {
			fmt.Printf("[DEBUG] Добавлена запись: студент='%s', дата='%s', дисциплина='%s'\n", currentStudent, date, discipline)
		}
	}

	// Обновляем общее количество студентов в группах
	totalStudents := 0
	for _, group := range groupsMap {
		// Маппим отделение по префиксу группы
		group.Department = departmentForGroup(group.Group)
		group.TotalStudents = len(group.Students)
		totalStudents += group.TotalStudents
		fmt.Printf("[DEBUG] Группа '%s': %d студентов\n", group.Group, group.TotalStudents)
	}

	// Преобразуем map в slice
	groups := make([]GroupLessons, 0, len(groupsMap))
	for _, g := range groupsMap {
		groups = append(groups, *g)
	}

	// Формируем итоговую структуру
	output := LessonsOutput{
		Period:        period,
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

	fmt.Printf(" Конвертация расписания занятий завершена.\n")
	fmt.Printf("   Период: %s\n", output.Period)
	fmt.Printf("   Групп: %d\n", output.TotalGroups)
	fmt.Printf("   Студентов: %d\n", output.TotalStudents)
	fmt.Printf("   Файл сохранён: %s\n", outputPath)
	return nil
}

// parseDate парсит дату из строки Excel и возвращает в формате для группировки по дням недели и парам
func parseDate(dateStr string) (time.Time, error) {
	// Пробуем разные форматы даты
	formats := []string{
		"02.01.2006 15:04:05",
		"02.01.2006 0:00:00",
		"02.01.2006",
		"2006-01-02 15:04:05",
		"2006-01-02",
	}

	for _, format := range formats {
		if t, err := time.Parse(format, dateStr); err == nil {
			return t, nil
		}
	}

	return time.Time{}, fmt.Errorf("не удалось распарсить дату: %s", dateStr)
}

// isGroupCode определяет, выглядит ли строка как код группы (1ис1, 3вб3 и т.п.)
func isGroupCode(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	// Группа не должна содержать пробелов
	if strings.Contains(s, " ") {
		return false
	}

	// Должна начинаться с цифры
	runes := []rune(s)
	if len(runes) == 0 || runes[0] < '0' || runes[0] > '9' {
		return false
	}

	// Должна содержать хотя бы одну букву (латинскую или кириллическую)
	hasLetter := false
	for _, r := range runes {
		if (r >= 'а' && r <= 'я') || (r >= 'А' && r <= 'Я') ||
			(r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') {
			hasLetter = true
			break
		}
	}

	return hasLetter
}

// departmentForGroup возвращает отделение по префиксу группы
// Маппинг можно дополнять по мере необходимости
func departmentForGroup(group string) string {
	g := strings.ToLower(strings.TrimSpace(group))
	if g == "" {
		return ""
	}

	switch {
	// Примеры маппинга. Ты можешь расширить этот список под свои группы.
	// Отделение креативных индустрий: а, д, м, р (анимация, дизайн, музыка, реклама)
	case strings.HasPrefix(g, "1а") || strings.HasPrefix(g, "2а") || strings.HasPrefix(g, "3а") ||
		strings.HasPrefix(g, "1д") || strings.HasPrefix(g, "2д") || strings.HasPrefix(g, "3д") ||
		strings.HasPrefix(g, "1м") || strings.HasPrefix(g, "2м") || strings.HasPrefix(g, "3м") ||
		strings.HasPrefix(g, "1р") || strings.HasPrefix(g, "2р") || strings.HasPrefix(g, "3р"):
		return "Отделение креативных индустрий"
	// Отделение программирования: вб, пк
	case strings.HasPrefix(g, "1ис") || strings.HasPrefix(g, "2ис") || strings.HasPrefix(g, "3ис"):
		return "Отделение программирования"
	case strings.Contains(g, "вб") || strings.Contains(g, "пк"):
		return "Отделение программирования"
	// Отделение ИТ и беспилотников: ис, са, бп
	case strings.HasPrefix(g, "1са") || strings.HasPrefix(g, "2са") || strings.HasPrefix(g, "3са") ||
		strings.HasPrefix(g, "1бп") || strings.HasPrefix(g, "2бп") || strings.HasPrefix(g, "3бп"):
		return "Отделение программирования"
	case strings.Contains(g, "са") || strings.Contains(g, "бп"):
		return "Отделение информационных технологий и беспилотников"
	// Отделение экономики: бд, бу
	case strings.HasPrefix(g, "1бд") || strings.HasPrefix(g, "2бд") || strings.HasPrefix(g, "3бд"):
		return "Отделение экономики"
	case strings.Contains(g, "бд") || strings.Contains(g, "бу"):
		return "Отделение экономики"
	default:
		return "Неизвестное отделение"
	}
}
