package services

import (
	"fmt"
	"time"

	"dashboard/internal/data"
	"dashboard/internal/models"
)

// ScheduleService читает schedule.json и даёт удобные выборки по датам.
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

// GetPlannedForDate возвращает всех студентов, которые должны быть на занятиях в указанную дату.
// date ожидается в формате YYYY-MM-DD.
// Дедуплицирует студентов: если у студента несколько пар в день, он учитывается один раз.
func (s *ScheduleService) GetPlannedForDate(date string) ([]PlannedStudent, error) {
	if s.schedulePath == "" || s.jsonStore == nil {
		return nil, fmt.Errorf("schedule service is not configured with path")
	}

	schedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, s.schedulePath)
	if err != nil {
		return nil, err
	}

	// schedule.json хранит даты в формате "02.02.2026 0:00:00"
	target, err := time.Parse("2006-01-02", date)
	if err != nil {
		return nil, fmt.Errorf("invalid date format, expected YYYY-MM-DD: %w", err)
	}
	targetDay := target.Format("02.01.2006")

	// Используем map для дедупликации: ключ = Department|Group|Student
	seen := make(map[string]bool)
	var out []PlannedStudent

	for _, g := range schedule.Groups {
		for _, st := range g.Students {
			key := g.Department + "|" + g.Group + "|" + st.StudentName
			// Проверяем, есть ли у студента хотя бы одна запись на эту дату
			hasRecord := false
			var firstRec models.ScheduleLessonRecord
			for _, rec := range st.Records {
				if rec.Date == "" {
					continue
				}
				if !hasDatePrefix(rec.Date, targetDay) {
					continue
				}
				hasRecord = true
				firstRec = rec
				break // Достаточно одной записи для подтверждения присутствия
			}
			if hasRecord && !seen[key] {
				seen[key] = true
				out = append(out, PlannedStudent{
					Department: g.Department,
					Group:      g.Group,
					Student:    st.StudentName,
					Date:       date,
					Discipline: firstRec.Discipline,
					Teacher:    firstRec.Teacher,
				})
			}
		}
	}

	return out, nil
}

// GetPlannedForDateAndLesson возвращает студентов, запланированных на указанную дату и номер пары (1-6).
// Если lessonNumber == 0, возвращает всех на дату (как GetPlannedForDate).
// Дедуплицирует студентов: если у студента несколько записей на эту пару, он учитывается один раз.
func (s *ScheduleService) GetPlannedForDateAndLesson(date string, lessonNumber int) ([]PlannedStudent, error) {
	if s.schedulePath == "" || s.jsonStore == nil {
		return nil, fmt.Errorf("schedule service is not configured with path")
	}

	schedule, err := data.LoadJSON[models.ScheduleJSON](s.jsonStore, s.schedulePath)
	if err != nil {
		return nil, err
	}

	target, err := time.Parse("2006-01-02", date)
	if err != nil {
		return nil, fmt.Errorf("invalid date format, expected YYYY-MM-DD: %w", err)
	}
	targetDay := target.Format("02.01.2006")

	// Используем map для дедупликации: ключ = Department|Group|Student
	seen := make(map[string]bool)
	var out []PlannedStudent

	for _, g := range schedule.Groups {
		for _, st := range g.Students {
			key := g.Department + "|" + g.Group + "|" + st.StudentName
			// Проверяем, есть ли у студента запись на эту дату и пару
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
				break // Нашли подходящую запись
			}
			if matchingRec != nil && !seen[key] {
				seen[key] = true
				out = append(out, PlannedStudent{
					Department: g.Department,
					Group:      g.Group,
					Student:    st.StudentName,
					Date:       date,
					Discipline: matchingRec.Discipline,
					Teacher:    matchingRec.Teacher,
				})
			}
		}
	}

	return out, nil
}

func hasDatePrefix(raw, day string) bool {
	if len(raw) < len(day) {
		return false
	}
	return raw[:len(day)] == day
}

