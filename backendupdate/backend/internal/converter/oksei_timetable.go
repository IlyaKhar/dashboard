package converter

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/extrame/xls"
)

// OkseiTimetableItem — одна пара в расписании (группа, день, номер пары, дисциплина, преподаватель)
type OkseiTimetableItem struct {
	Group      string `json:"group"`
	DayOfWeek  int    `json:"dayOfWeek"`  // 1=Пн, 2=Вт, 3=Ср, 4=Чт, 5=Пт, 6=Сб
	PairNumber int    `json:"pairNumber"` // 1-6
	Discipline string `json:"discipline"`
	Teacher    string `json:"teacher"`
}

// OkseiGroupSchedule — расписание группы (день -> пара -> дисциплина/преподаватель)
type OkseiGroupSchedule struct {
	Group      string                      `json:"group"`
	Department string                      `json:"department"`
	Schedule   map[int]map[int]OkseiLesson `json:"schedule"` // dayOfWeek -> pairNumber -> lesson
}

// OkseiLesson — занятие (дисциплина + преподаватель)
type OkseiLesson struct {
	Discipline string `json:"discipline"`
	Teacher    string `json:"teacher"`
}

// OkseiTimetableOutput — итоговое расписание пар с сайта ОКЭИ
type OkseiTimetableOutput struct {
	Period   string               `json:"period"` // например "09.02 - 14.02"
	Groups   []OkseiGroupSchedule `json:"groups"`
	RawItems []OkseiTimetableItem `json:"-"` // для обратной совместимости
}

var (
	dayNames = []string{"", "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"}
	dayRe    = regexp.MustCompile(`(?i)(Понедельник|Вторник|Среда|Четверг|Пятница|Суббота)`)
	pairRe   = regexp.MustCompile(`^[1-6]$`)
	lessonRe = regexp.MustCompile(`^(\d+)\.\s*(.+)`) // "1. Дисциплина Преподаватель"
	rovRe    = regexp.MustCompile(`(?i)^\s*РОВ\s+(.+)$`)
)

// ParseOkseiTimetable парсит .xls расписание ОКЭИ и возвращает структуру
func ParseOkseiTimetable(inputPath string) (*OkseiTimetableOutput, error) {
	wb, err := xls.Open(inputPath, "utf-8")
	if err != nil {
		return nil, fmt.Errorf("открытие файла: %w", err)
	}

	sheet := wb.GetSheet(0)
	if sheet == nil {
		return nil, fmt.Errorf("лист не найден")
	}

	maxRow := int(sheet.MaxRow)
	if maxRow < 1 {
		return nil, fmt.Errorf("файл пуст")
	}

	// Строка 0: заголовки — группы в колонках 2+
	header := sheet.Row(0)
	var groups []string
	for ci := 2; ci < 200; ci++ {
		g := strings.TrimSpace(header.Col(ci))
		if g == "" || g == "День недели" || g == "Номер пары" {
			continue
		}
		if !isGroupCode(g) {
			break
		}
		groups = append(groups, g)
	}

	if len(groups) == 0 {
		return nil, fmt.Errorf("группы не найдены в заголовке")
	}

	period := extractPeriodFromFilename(inputPath)
	out := &OkseiTimetableOutput{Period: period}
	byGroup := make(map[string]*OkseiGroupSchedule)
	for _, g := range groups {
		byGroup[g] = &OkseiGroupSchedule{
			Group:      g,
			Department: departmentForGroup(g),
			Schedule:   make(map[int]map[int]OkseiLesson),
		}
	}

	var currentDay int
	var currentPair int

	for ri := 1; ri <= maxRow; ri++ {
		row := sheet.Row(ri)
		col0 := strings.TrimSpace(row.Col(0))
		col1 := strings.TrimSpace(row.Col(1))

		// Обновляем день недели
		if m := dayRe.FindString(col0); m != "" {
			for d, name := range dayNames {
				if strings.EqualFold(m, name) {
					currentDay = d
					break
				}
			}
		}
		// День может быть в col1 (Вторник и т.д. иногда попадают в колонку группы при сдвиге)
		if currentDay == 0 && dayRe.MatchString(col1) {
			for d, name := range dayNames {
				if strings.Contains(strings.ToLower(col1), strings.ToLower(name)) {
					currentDay = d
					break
				}
			}
		}

		// Номер пары
		if pairRe.MatchString(col1) {
			switch col1 {
			case "1":
				currentPair = 1
			case "2":
				currentPair = 2
			case "3":
				currentPair = 3
			case "4":
				currentPair = 4
			case "5":
				currentPair = 5
			case "6":
				currentPair = 6
			}
		}

		if currentDay == 0 || currentPair == 0 {
			continue
		}

		// Читаем ячейки групп (колонки 2+)
		for gi, groupName := range groups {
			colIdx := 2 + gi
			cell := row.Col(colIdx)
			cell = strings.TrimSpace(cell)
			if cell == "" {
				continue
			}

			disc, teacher := parseOkseiLessonCell(cell, currentPair)
			if disc == "" && teacher == "" {
				continue
			}

			gr := byGroup[groupName]
			if gr.Schedule[currentDay] == nil {
				gr.Schedule[currentDay] = make(map[int]OkseiLesson)
			}
			gr.Schedule[currentDay][currentPair] = OkseiLesson{Discipline: disc, Teacher: teacher}
		}
	}

	for _, g := range groups {
		out.Groups = append(out.Groups, *byGroup[g])
	}

	log.Printf("[OKSEI] Распарсено групп: %d, период: %s", len(out.Groups), out.Period)
	return out, nil
}

// ConvertOkseiTimetable парсит .xls и сохраняет JSON в outputPath
func ConvertOkseiTimetable(inputPath, outputPath string) error {
	out, err := ParseOkseiTimetable(inputPath)
	if err != nil {
		return err
	}

	absPath, err := filepath.Abs(outputPath)
	if err != nil {
		return fmt.Errorf("путь: %w", err)
	}

	if err := os.MkdirAll(filepath.Dir(absPath), 0o755); err != nil {
		return fmt.Errorf("создание папки: %w", err)
	}

	data, err := json.MarshalIndent(out, "", "  ")
	if err != nil {
		return fmt.Errorf("JSON: %w", err)
	}

	if err := os.WriteFile(absPath, data, 0o644); err != nil {
		return fmt.Errorf("запись файла: %w", err)
	}

	log.Printf("[OKSEI] Расписание сохранено в %s", absPath)
	return nil
}

func parseOkseiLessonCell(cell string, currentPair int) (discipline, teacher string) {
	lines := strings.Split(cell, "\n")
	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}
		if rovRe.MatchString(line) {
			sub := rovRe.FindStringSubmatch(line)
			if len(sub) >= 2 {
				teacher = strings.TrimSpace(sub[1])
				return "РОВ", teacher
			}
		}
		if lessonRe.MatchString(line) {
			sub := lessonRe.FindStringSubmatch(line)
			if len(sub) >= 3 {
				var pairNum int
				_, _ = fmt.Sscanf(sub[1], "%d", &pairNum)
				if pairNum != currentPair {
					continue
				}
				rest := strings.TrimSpace(sub[2])
				parts := strings.Fields(rest)
				if len(parts) == 0 {
					continue
				}
				// Последнее слово может быть аудиторией (число, "зал", "(-)" и т.д.) — отбрасываем
				end := len(parts)
				for end > 1 && isRoomNumber(parts[end-1]) {
					end--
				}
				if end < 1 {
					return rest, ""
				}
				parts = parts[:end]
				if len(parts) >= 2 {
					last := parts[len(parts)-1]
					// "И.О." в конце — берём последние 2 слова как преподаватель (Фамилия И.О.)
					if len([]rune(last)) <= 5 && strings.Contains(last, ".") {
						teacher = strings.Join(parts[len(parts)-2:], " ")
						discipline = strings.Join(parts[:len(parts)-2], " ")
					} else {
						teacher = parts[len(parts)-1]
						discipline = strings.Join(parts[:len(parts)-1], " ")
					}
					if discipline == "" {
						discipline = rest
						teacher = ""
					}
					return discipline, teacher
				}
				return rest, ""
			}
		}
	}
	return "", ""
}

func isRoomNumber(s string) bool {
	s = strings.Trim(s, "()")
	if s == "зал" || s == "-" || s == "" {
		return true
	}
	if _, err := strconv.Atoi(s); err == nil {
		return true
	}
	if len(s) <= 4 && regexp.MustCompile(`^\d+[а-яА-Я]?$`).MatchString(s) {
		return true
	}
	// Корпус-Аудитория, корпус-Спортивный и т.д.
	if strings.Contains(s, "корпус-") {
		return true
	}
	if s == "Производственный" || s == "Основной" {
		return true
	}
	return false
}

func extractPeriodFromFilename(path string) string {
	base := filepath.Base(path)
	base = strings.TrimSuffix(base, ".xls")
	base = strings.TrimSuffix(base, ".xlsx")
	return base
}

