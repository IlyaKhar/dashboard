package converter

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Типы данных для посещаемости

type AttendanceRecord struct {
	Date   string `json:"date"`
	Missed int    `json:"missed"`
}

type Student struct {
	Student    string            `json:"student"`
	Attendance []AttendanceRecord `json:"attendance"`
}

type Group struct {
	Group    string    `json:"group"`
	Students []Student `json:"students"`
}

type Department struct {
	Department string   `json:"department"`
	Groups     []Group  `json:"groups"`
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
					"department": currentDepartment,
					"group":      currentGroup,
					"student":    currentStudent,
					"date":       dateStr,
					"missed":     int(hoursValue),
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
			Date:   date,
			Missed: missed,
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

func parseDateValue(value string) string {
	value = strings.TrimSpace(value)
	if value == "" {
		return ""
	}

	if num, err := strconv.ParseFloat(value, 64); err == nil {
		if num >= 1 && num < 100000 {
			excelEpoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
			days := int(num)
			date := excelEpoch.AddDate(0, 0, days)
			return date.Format("2006-01-02")
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
