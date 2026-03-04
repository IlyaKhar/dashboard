package converter

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Типы данных для контингента студентов

type StudentInfo struct {
	SerialNumber  int    `json:"serialNumber"`  // № п/п
	NumberInGroup int    `json:"numberInGroup"` // № в группе
	FullName      string `json:"fullName"`      // ФИО студента
	Status        string `json:"status"`        // Статус (например, "Студент")
}

type GroupContingent struct {
	Group    string        `json:"group"`
	Students []StudentInfo `json:"students"`
}

type DepartmentContingent struct {
	Department    string            `json:"department"`
	Groups        []GroupContingent `json:"groups"`
	TotalStudents int               `json:"totalStudents"` // Общее количество студентов в отделении
}

type StudentsOutput struct {
	TotalStudents int                    `json:"totalStudents"` // Общее количество студентов во всех отделениях
	Departments   []DepartmentContingent `json:"departments"`
}

// ConvertStudents конвертирует файл "Контингент студентов" Excel в JSON
// inputFile - путь к файлу Ведомостьколва.xlsx (или другому файлу со списком студентов)
// outputFile - путь к выходному JSON файлу
func ConvertStudents(inputFile, outputFile string) error {
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

	departmentsMap := make(map[string]*DepartmentContingent)

	var currentDepartment string
	var currentGroup string

	// Пропускаем заголовки и ищем начало данных
	// Обычно заголовки: "Контингент студентов", "Параметры:", "Дата отчета:", "Статус студента:"
	// Затем заголовки таблицы: "Отделение", "Группа", "№ п/п", "№ в группе", "Студент"

	dataStartRow := -1
	headerRow := -1

	// Ищем строку с заголовками таблицы
	for i, row := range rows {
		if len(row) == 0 {
			continue
		}

		// Ищем строку с заголовками "Отделение", "Группа", "№ п/п", "№ в группе", "Студент"
		rowText := strings.Join(row, " ")
		if strings.Contains(rowText, "Отделение") &&
			strings.Contains(rowText, "Группа") &&
			(strings.Contains(rowText, "№ п/п") || strings.Contains(rowText, "п/п")) {
			headerRow = i
			dataStartRow = i + 1
			break
		}
	}

	if dataStartRow == -1 {
		// Если не нашли заголовки, начинаем с 3-й строки (после заголовка и параметров)
		dataStartRow = 3
	}

	// Определяем индексы колонок
	// Структура: A - отделение/группа, B/C/D - номер в группе (может быть в любой из этих колонок), E - ФИО студента
	var colDepartmentOrGroup, colNumberInGroup, colStudent int = -1, -1, -1

	if headerRow >= 0 && headerRow < len(rows) {
		headerRowData := rows[headerRow]
		for i, cell := range headerRowData {
			cellLower := strings.ToLower(strings.TrimSpace(cell))
			if strings.Contains(cellLower, "отделение") || strings.Contains(cellLower, "группа") {
				colDepartmentOrGroup = i
			} else if strings.Contains(cellLower, "№ в группе") || strings.Contains(cellLower, "в группе") {
				colNumberInGroup = i
			} else if strings.Contains(cellLower, "студент") {
				colStudent = i
			}
		}
	}

	// Если не нашли колонки в заголовках, используем дефолтные индексы
	// A (0) - отделение/группа, B/C/D (1/2/3) - номер в группе, E (4) - ФИО
	if colDepartmentOrGroup == -1 {
		colDepartmentOrGroup = 0 // Колонка A
	}
	if colNumberInGroup == -1 {
		// Номер в группе может быть в B, C или D (колонки 1, 2, 3)
		// Будем проверять все три колонки
		colNumberInGroup = 1 // Начинаем с колонки B
	}
	if colStudent == -1 {
		colStudent = 4 // Колонка E
	}

	// Обрабатываем данные
	for i := dataStartRow; i < len(rows); i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}

		// Пропускаем строки с заголовками или итогами
		firstCell := ""
		if len(row) > 0 {
			firstCell = strings.TrimSpace(row[0])
		}

		if firstCell == "" {
			continue
		}

		// Пропускаем служебные строки
		if strings.Contains(firstCell, "Контингент") ||
			strings.Contains(firstCell, "Параметры") ||
			strings.Contains(firstCell, "Дата отчета") ||
			strings.Contains(firstCell, "Статус студента") ||
			strings.Contains(firstCell, "Итого") ||
			firstCell == "Отделение" {
			continue
		}

		// Читаем значение из колонки A (отделение или группа)
		colAValue := ""
		if colDepartmentOrGroup >= 0 && colDepartmentOrGroup < len(row) {
			colAValue = strings.TrimSpace(row[colDepartmentOrGroup])
		}

		// Читаем номер в группе (проверяем колонки B, C, D - 1, 2, 3)
		numberInGroupStr := ""
		for colIdx := 1; colIdx <= 3 && colIdx < len(row); colIdx++ {
			cellValue := strings.TrimSpace(row[colIdx])
			if cellValue != "" {
				// Проверяем, является ли это числом
				if num, err := strconv.Atoi(cellValue); err == nil && num > 0 {
					numberInGroupStr = cellValue
					break
				}
			}
		}

		// Читаем ФИО студента из колонки E (индекс 4)
		studentName := ""
		if colStudent >= 0 && colStudent < len(row) {
			studentName = strings.TrimSpace(row[colStudent])
		}

		// Определяем, что это отделение
		if colAValue != "" && strings.HasPrefix(colAValue, "Отделение") {
			currentDepartment = colAValue
			currentGroup = ""

			if _, exists := departmentsMap[currentDepartment]; !exists {
				departmentsMap[currentDepartment] = &DepartmentContingent{
					Department:    currentDepartment,
					Groups:        []GroupContingent{},
					TotalStudents: 0,
				}
			}
			continue
		}

		// Определяем, что это группа (начинается с цифры, содержит буквы)
		// Если это не отделение и не пусто - проверяем, группа ли это
		group := ""
		if colAValue != "" && len(colAValue) <= 15 && !strings.HasPrefix(colAValue, "Отделение") {
			firstChar := colAValue[0]
			if firstChar >= '0' && firstChar <= '9' {
				hasLetters := false
				for _, r := range colAValue {
					if (r >= 'а' && r <= 'я') || (r >= 'А' && r <= 'Я') ||
						(r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') {
						hasLetters = true
						break
					}
				}
				if hasLetters {
					group = colAValue
				}
			}
		}

		// Если это группа (есть название группы, но нет студента) - обновляем текущую группу
		if group != "" && numberInGroupStr == "" && studentName == "" {
			currentGroup = group
			if currentDepartment != "" {
				dept := departmentsMap[currentDepartment]
				var groupObj *GroupContingent
				for i := range dept.Groups {
					if dept.Groups[i].Group == currentGroup {
						groupObj = &dept.Groups[i]
						break
					}
				}
				if groupObj == nil {
					dept.Groups = append(dept.Groups, GroupContingent{
						Group:    currentGroup,
						Students: []StudentInfo{},
					})
				}
			}
			continue
		}

		// Если есть группа в строке (даже если она уже в контексте) - обновляем текущую группу
		if group != "" {
			currentGroup = group
			if currentDepartment != "" {
				dept := departmentsMap[currentDepartment]
				var groupObj *GroupContingent
				for i := range dept.Groups {
					if dept.Groups[i].Group == currentGroup {
						groupObj = &dept.Groups[i]
						break
					}
				}
				if groupObj == nil {
					dept.Groups = append(dept.Groups, GroupContingent{
						Group:    currentGroup,
						Students: []StudentInfo{},
					})
				}
			}
		}

		// Определяем, что это группа (а не студент)
		// Критерии для группы:
		// 1. В колонке "Группа" есть значение, начинающееся с цифры и содержащее буквы (например, "16п1")
		// 2. В колонке "Студент" НЕТ ФИО (пусто или меньше 2 слов)
		// 3. В колонке "№ в группе" НЕТ числа
		isGroup := false
		if group != "" && len(group) <= 15 {
			// Проверяем, что группа начинается с цифры и содержит буквы
			firstChar := group[0]
			if firstChar >= '0' && firstChar <= '9' {
				hasLetters := false
				for _, r := range group {
					if (r >= 'а' && r <= 'я') || (r >= 'А' && r <= 'Я') ||
						(r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') {
						hasLetters = true
						break
					}
				}

				if hasLetters {
					// Проверяем наличие ФИО студента
					words := strings.Fields(studentName)
					hasStudentName := len(words) >= 2

					// Проверяем наличие номера в группе
					hasNumberInGroup := false
					if numberInGroupStr != "" {
						if _, err := strconv.Atoi(numberInGroupStr); err == nil {
							hasNumberInGroup = true
						}
					}

					// Это группа, если: есть название группы, НО нет ФИО студента И нет номера в группе
					if !hasStudentName && !hasNumberInGroup {
						isGroup = true
					}
				}
			}
		}

		// Если это группа - обновляем текущую
		if isGroup && group != "" {
			// Сохраняем группу как есть (не в нижнем регистре!)
			currentGroup = group

			// Создаём группу, если её нет
			if currentDepartment != "" {
				dept := departmentsMap[currentDepartment]
				var groupObj *GroupContingent
				for i := range dept.Groups {
					if dept.Groups[i].Group == currentGroup {
						groupObj = &dept.Groups[i]
						break
					}
				}
				if groupObj == nil {
					dept.Groups = append(dept.Groups, GroupContingent{
						Group:    currentGroup,
						Students: []StudentInfo{},
					})
				}
			}
			continue
		}

		// Если нет отделения - пропускаем
		if currentDepartment == "" {
			continue
		}

		// Проверяем, есть ли номер в группе - это признак студента
		hasNumberInGroup := false
		if numberInGroupStr != "" {
			if _, err := strconv.Atoi(numberInGroupStr); err == nil {
				hasNumberInGroup = true
			}
		}

		// Проверяем, есть ли ФИО студента (минимум 2 слова)
		words := strings.Fields(studentName)
		hasStudentName := len(words) >= 2

		// Если есть группа в строке (даже если она уже в контексте) - обновляем текущую группу
		if group != "" && len(group) <= 15 {
			firstChar := group[0]
			if firstChar >= '0' && firstChar <= '9' {
				hasLetters := false
				for _, r := range group {
					if (r >= 'а' && r <= 'я') || (r >= 'А' && r <= 'Я') ||
						(r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') {
						hasLetters = true
						break
					}
				}
				if hasLetters {
					currentGroup = group
					// Создаём группу, если её нет
					if currentDepartment != "" {
						dept := departmentsMap[currentDepartment]
						var groupObj *GroupContingent
						for i := range dept.Groups {
							if dept.Groups[i].Group == currentGroup {
								groupObj = &dept.Groups[i]
								break
							}
						}
						if groupObj == nil {
							dept.Groups = append(dept.Groups, GroupContingent{
								Group:    currentGroup,
								Students: []StudentInfo{},
							})
						}
					}
				}
			}
		}

		// Если нет группы - пропускаем
		if currentGroup == "" {
			continue
		}

		// Если нет имени студента И нет номера в группе - пропускаем (это была строка группы)
		if !hasStudentName && !hasNumberInGroup {
			continue
		}

		// Если есть номер в группе, но нет имени - тоже пропускаем (неполные данные)
		if hasNumberInGroup && !hasStudentName {
			continue
		}

		// Парсим номер в группе
		numberInGroup := 0
		if numberInGroupStr != "" {
			if num, err := strconv.Atoi(numberInGroupStr); err == nil {
				numberInGroup = num
			}
		}

		// Добавляем студента
		dept := departmentsMap[currentDepartment]
		var groupObj *GroupContingent
		for i := range dept.Groups {
			if dept.Groups[i].Group == currentGroup {
				groupObj = &dept.Groups[i]
				break
			}
		}

		if groupObj == nil {
			dept.Groups = append(dept.Groups, GroupContingent{
				Group:    currentGroup,
				Students: []StudentInfo{},
			})
			groupObj = &dept.Groups[len(dept.Groups)-1]
		}

		// Проверяем, нет ли уже такого студента
		exists := false
		for _, s := range groupObj.Students {
			if s.FullName == studentName {
				exists = true
				break
			}
		}

		if !exists {
			groupObj.Students = append(groupObj.Students, StudentInfo{
				SerialNumber:  numberInGroup, // Используем номер в группе как серийный номер
				NumberInGroup: numberInGroup,
				FullName:      studentName,
				Status:        "Студент", // По умолчанию, можно брать из параметров
			})
		}
	}

	// Подсчитываем общее количество студентов и по отделениям
	totalStudents := 0
	maxNumberInGroup := 0 // Последний номер в группе = общее количество

	// Преобразуем map в slice и считаем студентов
	departments := make([]DepartmentContingent, 0, len(departmentsMap))
	for _, d := range departmentsMap {
		deptTotal := 0
		for _, g := range d.Groups {
			deptTotal += len(g.Students)
			totalStudents += len(g.Students)
			// Находим максимальный номер в группе (последний номер = общее количество)
			for _, s := range g.Students {
				if s.NumberInGroup > maxNumberInGroup {
					maxNumberInGroup = s.NumberInGroup
				}
			}
		}
		d.TotalStudents = deptTotal
		departments = append(departments, *d)
	}

	// Если нашли максимальный номер в группе, используем его как общее количество
	// Иначе используем фактическое количество студентов
	if maxNumberInGroup > totalStudents {
		totalStudents = maxNumberInGroup
	}

	// Формируем итоговую структуру
	output := StudentsOutput{
		TotalStudents: totalStudents,
		Departments:   departments,
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

	fmt.Printf(" Конвертация контингента студентов завершена.\n")
	fmt.Printf("   Общее количество студентов: %d\n", output.TotalStudents)
	fmt.Printf("   Отделений: %d\n", len(output.Departments))
	totalGroups := 0
	for _, d := range output.Departments {
		totalGroups += len(d.Groups)
		fmt.Printf("   - %s: %d студентов, %d групп\n", d.Department, d.TotalStudents, len(d.Groups))
	}
	fmt.Printf("   Всего групп: %d\n", totalGroups)
	fmt.Printf("   Файл сохранён: %s\n", outputPath)
	return nil
}
