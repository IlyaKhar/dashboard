package parse

import (
	"strconv"
	"strings"
	"time"
)

// Форматы дат для парсинга
var dateFormats = []string{
	"02.01.2006 15:04:05",
	"02.01.2006",
	"02/01/2006",
	"2006-01-02",
	"02.01.06",
	"02/01/06",
}

// ParseDate парсит дату из строки или Excel серийного номера
func ParseDate(value string) string {
	value = strings.TrimSpace(value)
	if value == "" {
		return ""
	}

	// Пробуем парсить как Excel серийный номер (число)
	if num, err := strconv.ParseFloat(value, 64); err == nil {
		if num >= 1 && num < 100000 {
			excelEpoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
			days := int(num)
			date := excelEpoch.AddDate(0, 0, days)
			return date.Format("2006-01-02")
		}
	}

	// Пробуем парсить как строку в различных форматах
	for _, format := range dateFormats {
		if parsed, err := time.Parse(format, value); err == nil {
			return parsed.Format("2006-01-02")
		}
	}
	return ""
}
