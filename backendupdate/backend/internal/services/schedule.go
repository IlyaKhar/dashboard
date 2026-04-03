package services

import (
	"fmt"
	"strings"
	"time"

	"dashboard/internal/data"
	"dashboard/internal/models"
)

// ScheduleService читает schedule.json (сегодня) и schedule_history.json (история).
// Для поиска по дате сначала пробует историю, потом текущий файл.
type ScheduleService struct {
	schedulePath string
	jsonStore    *data.JSONStore
}

func NewScheduleService(schedulePath string) *ScheduleService {
	return &ScheduleService{
		schedulePath: schedulePath,
		jsonStore:    data.NewJSONStore(),
	}
}

// PlannedStudent описывает запланированное присутствие студента на паре.
type PlannedStudent struct {
	Department string
	Group      string
	Student    string
	Date       string
	Discipline string
	Teacher    string
}

// historyPath возвращает путь к schedule_history.json на основе пути к schedule.json
func (s *ScheduleService) historyPath() string {
	return strings.Replace(s.schedulePath, "schedule.json", "schedule_history.json", 1)
}

// GetPlannedForDate возвращает всех студентов, которые должны быть на занятиях в указанную дату.
// date ожидается в формате YYYY-MM-DD.
// Сначала ищет в schedule_history.json (накопленная история), затем в schedule.json (сегодня).
// Дедуплицирует студентов: если у студента несколько пар в день, он учитывается один раз.
func (s *ScheduleService) GetPlannedForDate(date string) ([]PlannedStudent, error) {
	if s.schedulePath == "" || s.jsonStore == nil {
		return nil, fmt.Errorf("schedule service is not configured with path")
	}

	target, err := time.Parse("2006-01-02", date)
	if err != nil {
		return nil, fmt.Errorf("invalid date format, expected YYYY-MM-DD: %w", err)
	}
	targetDay := target.Format("02.01.2006")

	// Сначала ищем в истории
	histPath := s.historyPath()
	histSchedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, histPath)
	if err == nil {
		out := extractPlannedForDate(histSchedule, targetDay, date, 0)
		if len(out) > 0 {
			return out, nil
		}
	}

	// Если в истории не нашли — ищем в текущем schedule.json
	schedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, s.schedulePath)
	if err != nil {
		return nil, err
	}

	return extractPlannedForDate(schedule, targetDay, date, 0), nil
}

// GetPlannedForDateAndLesson возвращает студентов, запланированных на указанную дату и номер пары (1-6).
// Сначала ищет в schedule_history.json, затем в schedule.json.
func (s *ScheduleService) GetPlannedForDateAndLesson(date string, lessonNumber int) ([]PlannedStudent, error) {
	if s.schedulePath == "" || s.jsonStore == nil {
		return nil, fmt.Errorf("schedule service is not configured with path")
	}

	target, err := time.Parse("2006-01-02", date)
	if err != nil {
		return nil, fmt.Errorf("invalid date format, expected YYYY-MM-DD: %w", err)
	}
	targetDay := target.Format("02.01.2006")

	// Сначала ищем в истории
	histPath := s.historyPath()
	histSchedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, histPath)
	if err == nil {
		out := extractPlannedForDate(histSchedule, targetDay, date, lessonNumber)
		if len(out) > 0 {
			return out, nil
		}
	}

	// Если в истории не нашли — ищем в текущем schedule.json
	schedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, s.schedulePath)
	if err != nil {
		return nil, err
	}

	return extractPlannedForDate(schedule, targetDay, date, lessonNumber), nil
}

// extractPlannedForDate извлекает запланированных студентов из schedule за конкретную дату.
// Если lessonNumber > 0, фильтрует по номеру пары.
func extractPlannedForDate(schedule models.ScheduleJSON, targetDay, dateISO string, lessonNumber int) []PlannedStudent {
	seen := make(map[string]bool)
	var out []PlannedStudent

	for _, g := range schedule.Groups {
		for _, st := range g.Students {
			key := g.Department + "|" + g.Group + "|" + st.StudentName
			var matchingRec *models.ScheduleLessonRecord
			for _, rec := range st.Records {
				if rec.Date == "" {
					continue
				}
				if !hasDatePrefix(rec.Date, targetDay) {
					continue
				}
				if lessonNumber > 0 && rec.LessonNumber != lessonNumber {
					continue
				}
				matchingRec = &rec
				break
			}
			if matchingRec != nil && !seen[key] {
				seen[key] = true
				out = append(out, PlannedStudent{
					Department: g.Department,
					Group:      g.Group,
					Student:    st.StudentName,
					Date:       dateISO,
					Discipline: matchingRec.Discipline,
					Teacher:    matchingRec.Teacher,
				})
			}
		}
	}

	return out
}

func hasDatePrefix(raw, day string) bool {
	if len(raw) < len(day) {
		return false
	}
	return raw[:len(day)] == day
}

// GetGroupLessonMinForDate возвращает минимальный номер пары по каждой группе на дату.
// Ключ карты: "department|group".
func (s *ScheduleService) GetGroupLessonMinForDate(date string) (map[string]int, error) {
	if s.schedulePath == "" || s.jsonStore == nil {
		return nil, fmt.Errorf("schedule service is not configured with path")
	}

	target, err := time.Parse("2006-01-02", date)
	if err != nil {
		return nil, fmt.Errorf("invalid date format, expected YYYY-MM-DD: %w", err)
	}
	targetDay := target.Format("02.01.2006")

	// Сначала пробуем историю.
	histPath := s.historyPath()
	histSchedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, histPath)
	if err == nil {
		minMap := extractGroupLessonMin(histSchedule, targetDay)
		if len(minMap) > 0 {
			return minMap, nil
		}
	}

	// Если в истории нет нужной даты — читаем текущий файл.
	schedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, s.schedulePath)
	if err != nil {
		return nil, err
	}
	return extractGroupLessonMin(schedule, targetDay), nil
}

func extractGroupLessonMin(schedule models.ScheduleJSON, targetDay string) map[string]int {
	out := make(map[string]int)
	for _, g := range schedule.Groups {
		key := g.Department + "|" + g.Group
		for _, st := range g.Students {
			for _, rec := range st.Records {
				if rec.Date == "" || !hasDatePrefix(rec.Date, targetDay) || rec.LessonNumber <= 0 {
					continue
				}
				if prev, ok := out[key]; !ok || rec.LessonNumber < prev {
					out[key] = rec.LessonNumber
				}
			}
		}
	}
	return out
}
