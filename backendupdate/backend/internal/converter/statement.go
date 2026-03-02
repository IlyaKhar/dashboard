package converter

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

// Типы данных для ведомости

type StudentSummary struct {
	Student       string `json:"student"`
	MissedTotal   int    `json:"missedTotal"`
	MissedBad     int    `json:"missedBad"`     // не по уважительной
	MissedExcused int    `json:"missedExcused"` // по уважительной
}

type DisciplineSummary struct {
	Discipline  string           `json:"discipline"`
	TotalMissed int              `json:"totalMissed"`
	Students    []StudentSummary `json:"students"`
}

type GroupSummary struct {
	Group       string              `json:"group"`
	TotalMissed int                 `json:"totalMissed"`
	Disciplines []DisciplineSummary `json:"disciplines,omitempty"` // Опционально: если есть дисциплины
	Students    []StudentSummary    `json:"students,omitempty"`     // Опционально: если нет дисциплин
}

type SpecialtySummary struct {
	Specialty   string         `json:"specialty"`
	TotalMissed int            `json:"totalMissed"`
	Groups      []GroupSummary `json:"groups"`
}

type DepartmentSummary struct {
	Department  string             `json:"department"`
	TotalMissed int                `json:"totalMissed"`
	Specialties []SpecialtySummary `json:"specialties"`
}

// ConvertStatement конвертирует файл ведомости Excel в JSON
// inputFileXLS - путь к файлу ведомость.xls (или .xlsx)
// outputFile - путь к выходному JSON файлу
// pythonScriptPath - путь к Python скрипту для конвертации XLS → XLSX
func ConvertStatement(inputFileXLS, outputFile, pythonScriptPath string) error {
	// Определяем имя XLSX файла
	inputFileXLSX := strings.TrimSuffix(inputFileXLS, ".xls") + ".xlsx"

	// Шаг 1: Конвертируем XLS в XLSX (если нужно)
	if strings.HasSuffix(strings.ToLower(inputFileXLS), ".xls") {
		if _, err := os.Stat(inputFileXLS); err == nil {
			if err := convertXLSToXLSX(inputFileXLS, inputFileXLSX, pythonScriptPath); err != nil {
				fmt.Printf("Предупреждение при конвертации %s: %v\n", inputFileXLS, err)
				fmt.Println("Продолжаем с XLSX файлом, если он существует...")
			}
		}
	} else {
		// Если уже XLSX, используем его напрямую
		inputFileXLSX = inputFileXLS
	}

	// Шаг 2: Открываем XLSX файл через excelize
	f, err := excelize.OpenFile(inputFileXLSX)
	if err != nil {
		return fmt.Errorf("ошибка открытия файла %s: %v\nУбедитесь, что файл конвертирован в XLSX формат", inputFileXLSX, err)
	}
	defer f.Close()

	// Берём первый лист
	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return fmt.Errorf("не найден лист в файле")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("ошибка чтения строк: %v", err)
	}

	departmentsMap := make(map[string]*DepartmentSummary)

	var currentDepartment string
	var currentSpecialty string
	var currentGroup string
	var currentDiscipline string
	var period string

	// Перебираем все строки листа
	for _, row := range rows {
		if len(row) == 0 {
			continue
		}

		label := strings.TrimSpace(row[0])
		if label == "" {
			continue
		}

		// Извлекаем период из строки "Параметры:" (в той же строке может быть "Период: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ")
		if label == "Параметры:" {
			if p := extractPeriodFromRow(row); p != "" {
				period = p
			}
			continue
		}

		if isHeaderOrTotal(label) {
			continue
		}

		// Берём числа из колонок
		// Структура файла: A (текст), B-C (пусто), D (неуваж, обычно пусто), E (уваж), F-G (пусто), H (всего)
		bad := 0
		excused := 0
		total := 0

		n := len(row)

		// Колонка E (индекс 4) - пропущено по уважительной причине
		if n > 4 {
			excused = parseIntCell(row[4])
		}

		// Колонка H (индекс 7) - всего пропущено часов
		if n > 7 {
			total = parseIntCell(row[7])
		} else if n > 4 {
			// Если колонки H нет, используем E как total
			total = excused
		}

		// Колонка D (индекс 3) - пропущено не по уважительной причине (обычно пусто)
		if n > 3 {
			bad = parseIntCell(row[3])
		}

		// Если total = 0, но есть excused, используем excused как total
		if total == 0 && excused > 0 {
			total = excused
		}

		// Классифицируем строку
		if isDepartment(label) {
			currentDepartment = label
			currentSpecialty = ""
			currentGroup = ""
			currentDiscipline = ""

			if _, ok := departmentsMap[currentDepartment]; !ok {
				departmentsMap[currentDepartment] = &DepartmentSummary{
					Department:  currentDepartment,
					TotalMissed: 0,
					Specialties: []SpecialtySummary{},
				}
			}
			// Обновляем totalMissed для отделения, если есть числа
			if total > 0 {
				departmentsMap[currentDepartment].TotalMissed = total
			}
			continue
		}

		if isSpecialty(label) {
			currentSpecialty = label
			currentGroup = ""
			currentDiscipline = ""
			// Обновляем totalMissed для специальности, если есть числа
			if total > 0 && currentDepartment != "" {
				dept := departmentsMap[currentDepartment]
				var spec *SpecialtySummary
				for i := range dept.Specialties {
					if dept.Specialties[i].Specialty == currentSpecialty {
						spec = &dept.Specialties[i]
						break
					}
				}
				if spec == nil {
					dept.Specialties = append(dept.Specialties, SpecialtySummary{
						Specialty:   currentSpecialty,
						TotalMissed: total,
						Groups:      []GroupSummary{},
					})
				} else {
					spec.TotalMissed = total
				}
			}
			continue
		}

		if isGroup(label) {
			currentGroup = strings.ToLower(label)
			currentDiscipline = ""
			// Обновляем totalMissed для группы, если есть числа
			if total > 0 && currentDepartment != "" && currentSpecialty != "" {
				dept := departmentsMap[currentDepartment]
				var spec *SpecialtySummary
				for i := range dept.Specialties {
					if dept.Specialties[i].Specialty == currentSpecialty {
						spec = &dept.Specialties[i]
						break
					}
				}
				if spec != nil {
					var group *GroupSummary
					for i := range spec.Groups {
						if spec.Groups[i].Group == currentGroup {
							group = &spec.Groups[i]
							break
						}
					}
					if group == nil {
						spec.Groups = append(spec.Groups, GroupSummary{
							Group:       currentGroup,
							TotalMissed: total,
							Students:    []StudentSummary{},
							Disciplines: []DisciplineSummary{},
						})
					} else {
						group.TotalMissed = total
					}
				}
			}
			continue
		}

		// Проверяем, является ли строка дисциплиной
		// Дисциплина обычно длиннее названия группы и не является ФИО студента
		if isDiscipline(label) && currentDepartment != "" && currentSpecialty != "" && currentGroup != "" {
			currentDiscipline = label
			// Создаём дисциплину в группе, если её нет
			dept := departmentsMap[currentDepartment]
			var spec *SpecialtySummary
			for i := range dept.Specialties {
				if dept.Specialties[i].Specialty == currentSpecialty {
					spec = &dept.Specialties[i]
					break
				}
			}
			if spec != nil {
				var group *GroupSummary
				for i := range spec.Groups {
					if spec.Groups[i].Group == currentGroup {
						group = &spec.Groups[i]
						break
					}
				}
				if group != nil {
					// Проверяем, есть ли уже такая дисциплина
					hasDiscipline := false
					for i := range group.Disciplines {
						if group.Disciplines[i].Discipline == currentDiscipline {
							group.Disciplines[i].TotalMissed = total
							hasDiscipline = true
							break
						}
					}
					if !hasDiscipline {
						group.Disciplines = append(group.Disciplines, DisciplineSummary{
							Discipline:  currentDiscipline,
							TotalMissed: total,
							Students:    []StudentSummary{},
						})
					}
				}
			}
			continue
		}

		// Остальное считаем строками со студентами
		if currentDepartment == "" || currentSpecialty == "" || currentGroup == "" {
			continue
		}

		if total == 0 && bad == 0 && excused == 0 {
			continue
		}

		dept := departmentsMap[currentDepartment]

		// Ищем / создаём специальность
		var spec *SpecialtySummary
		for i := range dept.Specialties {
			if dept.Specialties[i].Specialty == currentSpecialty {
				spec = &dept.Specialties[i]
				break
			}
		}
		if spec == nil {
			dept.Specialties = append(dept.Specialties, SpecialtySummary{
				Specialty:   currentSpecialty,
				TotalMissed: 0,
				Groups:      []GroupSummary{},
			})
			spec = &dept.Specialties[len(dept.Specialties)-1]
		}

		// Ищем / создаём группу
		var group *GroupSummary
		for i := range spec.Groups {
			if spec.Groups[i].Group == currentGroup {
				group = &spec.Groups[i]
				break
			}
		}
		if group == nil {
			spec.Groups = append(spec.Groups, GroupSummary{
				Group:       currentGroup,
				TotalMissed: 0,
				Students:    []StudentSummary{},
				Disciplines: []DisciplineSummary{},
			})
			group = &spec.Groups[len(spec.Groups)-1]
		}

		// Добавляем студента
		student := StudentSummary{
			Student:       label,
			MissedTotal:   total,
			MissedBad:     bad,
			MissedExcused: excused,
		}

		// Если есть текущая дисциплина, добавляем студента в дисциплину
		if currentDiscipline != "" {
			var disc *DisciplineSummary
			for i := range group.Disciplines {
				if group.Disciplines[i].Discipline == currentDiscipline {
					disc = &group.Disciplines[i]
					break
				}
			}
			if disc == nil {
				group.Disciplines = append(group.Disciplines, DisciplineSummary{
					Discipline:  currentDiscipline,
					TotalMissed: 0,
					Students:    []StudentSummary{},
				})
				disc = &group.Disciplines[len(group.Disciplines)-1]
			}
			disc.Students = append(disc.Students, student)
			disc.TotalMissed += total
		} else {
			// Если нет дисциплины, добавляем студента напрямую в группу
			group.Students = append(group.Students, student)
		}

		// Обновляем суммы
		group.TotalMissed += total
		spec.TotalMissed += total
		dept.TotalMissed += total
	}

	// Преобразуем map в slice
	departments := make([]DepartmentSummary, 0, len(departmentsMap))
	for _, d := range departmentsMap {
		departments = append(departments, *d)
	}

	outputPath, err := filepath.Abs(outputFile)
	if err != nil {
		return fmt.Errorf("ошибка получения пути: %v", err)
	}

	// Выгрузка с полем period (период из шапки ведомости)
	root := vedomostOutput{Period: period, Departments: departments}
	data, err := json.MarshalIndent(root, "", "  ")
	if err != nil {
		return fmt.Errorf("ошибка сериализации JSON: %v", err)
	}

	if err := os.WriteFile(outputPath, data, 0644); err != nil {
		return fmt.Errorf("ошибка записи файла: %v", err)
	}

	fmt.Printf(" Конвертация ведомости завершена. Отделений: %d\n", len(departments))
	fmt.Printf("   Файл сохранён: %s\n", outputPath)
	return nil
}

// convertXLSToXLSX конвертирует XLS файл в XLSX формат через Python скрипт
func convertXLSToXLSX(xlsFile, xlsxFile, pythonScriptPath string) error {
	// Проверяем, существует ли уже XLSX файл и он новее XLS
	if info, err := os.Stat(xlsxFile); err == nil {
		if xlsInfo, err2 := os.Stat(xlsFile); err2 == nil {
			if info.ModTime().After(xlsInfo.ModTime()) {
				// XLSX файл новее, конвертация не нужна
				return nil
			}
		}
	}

	// Используем Python скрипт для конвертации
	return convertXLSToXLSXPython(xlsFile, xlsxFile, pythonScriptPath)
}

// convertXLSToXLSXPython использует Python скрипт для конвертации
func convertXLSToXLSXPython(xlsFile, xlsxFile, pythonScriptPath string) error {
	// Проверяем наличие Python скрипта
	if _, err := os.Stat(pythonScriptPath); os.IsNotExist(err) {
		// Пробуем найти скрипт в других местах
		possiblePaths := []string{
			pythonScriptPath,
			filepath.Join(filepath.Dir(xlsFile), "xls_to_xlsx.py"),
			filepath.Join(filepath.Dir(filepath.Dir(xlsFile)), "xls_to_xlsx.py"),
			filepath.Join(filepath.Dir(filepath.Dir(filepath.Dir(xlsFile))), "backend", "internal", "converter", "xls_to_xlsx.py"),
		}
		
		found := false
		for _, path := range possiblePaths {
			if _, err := os.Stat(path); err == nil {
				pythonScriptPath = path
				found = true
				fmt.Printf("[DEBUG] Python скрипт найден: %s\n", pythonScriptPath)
				break
			}
		}
		
		if !found {
			return fmt.Errorf("Python скрипт не найден. Проверенные пути: %v", possiblePaths)
		}
	} else {
		fmt.Printf("[DEBUG] Используем Python скрипт: %s\n", pythonScriptPath)
	}

	// Определяем команду Python
	var pythonCmd string
	if _, err := exec.LookPath("py"); err == nil {
		pythonCmd = "py"
	} else if _, err := exec.LookPath("python"); err == nil {
		pythonCmd = "python"
	} else if _, err := exec.LookPath("python3"); err == nil {
		pythonCmd = "python3"
	} else {
		return fmt.Errorf("Python не найден в PATH. Установите Python или добавьте его в PATH")
	}

	// Шаг 1: Проверяем и устанавливаем зависимости Python
	fmt.Printf("[DEBUG] Проверка зависимостей Python...\n")
	checkCmd := exec.Command(pythonCmd, "-m", "pip", "list")
	if pythonCmd == "py" {
		checkCmd = exec.Command("py", "-3", "-m", "pip", "list")
	}
	checkCmd.Dir = filepath.Dir(xlsFile)
	checkOutput, _ := checkCmd.CombinedOutput()
	checkOutputStr := string(checkOutput)
	
	// Проверяем наличие xlrd и openpyxl
	needsXlrd := !strings.Contains(checkOutputStr, "xlrd")
	needsOpenpyxl := !strings.Contains(checkOutputStr, "openpyxl")
	
	if needsXlrd || needsOpenpyxl {
		fmt.Printf("[DEBUG] Установка зависимостей Python: xlrd openpyxl\n")
		installArgs := []string{"-m", "pip", "install", "--quiet", "--upgrade", "xlrd", "openpyxl"}
		if pythonCmd == "py" {
			installArgs = append([]string{"-3"}, installArgs...)
		}
		installCmd := exec.Command(pythonCmd, installArgs...)
		installCmd.Dir = filepath.Dir(xlsFile)
		installOutput, err := installCmd.CombinedOutput()
		if err != nil {
			fmt.Printf("[DEBUG] Предупреждение при установке зависимостей: %v\n", err)
			fmt.Printf("[DEBUG] Вывод: %s\n", string(installOutput))
			// Продолжаем выполнение, возможно зависимости уже установлены
		} else {
			fmt.Printf("[DEBUG] Зависимости установлены успешно\n")
		}
	}

	// Шаг 2: Запускаем Python скрипт для конвертации
	var cmd *exec.Cmd
	var pythonArgs []string
	
	if pythonCmd == "py" {
		pythonArgs = []string{"-3", pythonScriptPath, xlsFile, xlsxFile}
	} else {
		pythonArgs = []string{pythonScriptPath, xlsFile, xlsxFile}
	}
	
	cmd = exec.Command(pythonCmd, pythonArgs...)
	cmd.Dir = filepath.Dir(xlsFile)
	
	output, err := cmd.CombinedOutput()
	if err != nil {
		// Пробуем декодировать вывод с учётом кодировки Windows
		outputStr := string(output)
		// Если вывод содержит ошибку о модуле, предлагаем решение
		if strings.Contains(outputStr, "No module named") {
			return fmt.Errorf("ошибка: не установлены зависимости Python\nУстановите: %s -m pip install xlrd openpyxl\nОшибка: %v\nВывод: %s", pythonCmd, err, outputStr)
		}
		return fmt.Errorf("ошибка конвертации XLS → XLSX через Python (%s): %v\nВывод: %s\nСкрипт: %s\nКоманда: %s %v", pythonCmd, err, outputStr, pythonScriptPath, pythonCmd, pythonArgs)
	}
	fmt.Printf("   Конвертировано через Python (%s): %s → %s\n", pythonCmd, xlsFile, xlsxFile)
	return nil
}

// Вспомогательные функции

func parseIntCell(value string) int {
	value = strings.TrimSpace(value)
	if value == "" {
		return 0
	}
	v, err := strconv.Atoi(value)
	if err != nil {
		// Пробуем как float и конвертируем в int
		if f, err2 := strconv.ParseFloat(value, 64); err2 == nil {
			return int(f)
		}
		return 0
	}
	return v
}

func isDepartment(text string) bool {
	return strings.HasPrefix(text, "Отделение ")
}

func isHeaderOrTotal(text string) bool {
	switch text {
	case "Сводная ведомость по посещаемости", "Параметры:", "Отделение", "Специальность", "Учебная группа", "Студент", "Итого":
		return true
	default:
		return false
	}
}

// vedomostOutput — корневой объект vedomost.json (period + departments)
type vedomostOutput struct {
	Period      string               `json:"period"`
	Departments []DepartmentSummary  `json:"departments"`
}

// extractPeriodFromRow ищет в ячейках строки подстроку "Период:" и возвращает значение после неё (например "18.02.2026 - 18.02.2026")
func extractPeriodFromRow(row []string) string {
	const prefix = "Период:"
	for _, cell := range row {
		s := strings.TrimSpace(cell)
		if idx := strings.Index(s, prefix); idx >= 0 {
			val := strings.TrimSpace(s[idx+len(prefix):])
			if val != "" {
				return val
			}
		}
	}
	return ""
}

func isSpecialty(text string) bool {
	text = strings.TrimSpace(text)
	if len(text) < 8 {
		return false
	}
	r := []rune(text)
	if !unicode.IsDigit(r[0]) {
		return false
	}
	if !strings.Contains(text, " ") {
		return false
	}
	if strings.Count(text, ".") < 2 {
		return false
	}
	return true
}

func isGroup(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" || len(text) > 10 {
		return false
	}
	r := []rune(text)
	if !unicode.IsDigit(r[0]) {
		return false
	}
	// Группа — одно "слово", без пробелов
	if len(strings.Fields(text)) != 1 {
		return false
	}
	return true
}

// isDiscipline определяет, является ли строка названием дисциплины
// Дисциплина обычно длиннее группы и содержит несколько слов
func isDiscipline(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}
	// Дисциплина обычно длиннее 10 символов и содержит пробелы или точки
	if len(text) < 10 {
		return false
	}
	// Проверяем, что это не группа (не начинается с цифры и короткое)
	if len(text) <= 10 && unicode.IsDigit([]rune(text)[0]) {
		return false
	}
	// Дисциплина обычно содержит несколько слов или специальные символы (точки, тире)
	words := strings.Fields(text)
	if len(words) >= 2 {
		return true
	}
	// Или содержит точки/тире (например, "МДК 05.03")
	if strings.Contains(text, ".") || strings.Contains(text, "-") {
		return true
	}
	return false
}
