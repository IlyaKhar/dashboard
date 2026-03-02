package parse

import (
	"strconv"
	"strings"
	"unicode"
)

// RowKind определяет тип строки в Excel
type RowKind int

const (
	RowUnknown RowKind = iota
	RowDepartment
	RowGroup
	RowStudent
)

func (k RowKind) String() string {
	switch k {
	case RowDepartment:
		return "department"
	case RowGroup:
		return "group"
	case RowStudent:
		return "student"
	default:
		return "unknown"
	}
}

// ClassifyRow определяет тип строки по первой ячейке
func ClassifyRow(firstCell string) RowKind {
	firstCell = strings.TrimSpace(firstCell)
	if firstCell == "" {
		return RowUnknown
	}

	// Отделение начинается с "Отделение"
	if strings.HasPrefix(firstCell, "Отделение") {
		return RowDepartment
	}

	// Группа - короткая строка, начинается с цифры
	if len(firstCell) <= 15 && len(firstCell) >= 1 &&
		strings.Count(firstCell, ".") != 2 && strings.Count(firstCell, "-") != 2 {
		r := rune(firstCell[0])
		if unicode.IsDigit(r) {
			for _, c := range firstCell {
				if !unicode.IsDigit(c) && c != '.' && c != '/' && c != ' ' && c != '-' {
					return RowUnknown
				}
			}
			return RowGroup
		}
	}

	// Студент - 3 слова (Фамилия Имя Отчество)
	words := strings.Fields(firstCell)
	if len(words) == 3 {
		// Если первое слово не число - это ФИО
		if _, err := strconv.ParseFloat(words[0], 64); err != nil {
			return RowStudent
		}
	}

	return RowUnknown
}
