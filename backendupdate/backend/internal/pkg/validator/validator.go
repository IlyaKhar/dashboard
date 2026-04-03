package validator

import (
	"regexp"
	"strings"
)

const (
	MaxDepartmentLen = 200
	MaxGroupLen      = 100
	MaxStudentLen    = 255
	MaxDateLen       = 10
	DatePattern      = `^\d{4}-\d{2}-\d{2}$`
)

// SanitizeString обрезает до maxLen и убирает опасные символы
func SanitizeString(s string, maxLen int) string {
	s = strings.TrimSpace(s)
	if len(s) > maxLen {
		s = s[:maxLen]
	}
	return s
}

// ValidateDepartment проверяет параметр department
func ValidateDepartment(s string) (string, bool) {
	s = SanitizeString(s, MaxDepartmentLen)
	if s == "" {
		return s, true // пустое допустимо для фильтра
	}
	return s, len(s) <= MaxDepartmentLen
}

// ValidateGroup проверяет параметр group
func ValidateGroup(s string) (string, bool) {
	s = SanitizeString(s, MaxGroupLen)
	if s == "" {
		return s, true
	}
	return s, len(s) <= MaxGroupLen
}

// ValidateStudent проверяет параметр student
func ValidateStudent(s string) (string, bool) {
	s = SanitizeString(s, MaxStudentLen)
	if s == "" {
		return s, true
	}
	return s, len(s) <= MaxStudentLen
}

// ValidateDate проверяет формат даты YYYY-MM-DD
func ValidateDate(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return true
	}
	ok, _ := regexp.MatchString(DatePattern, s)
	return ok
}
