package services

import (
	"fmt"
	"sync"
	"time"
)

// ReconciliationService сверяет расписание (кто должен быть) и посещаемость (кто реально был).
type ReconciliationService struct {
	attendance *AttendanceService
	schedule   *ScheduleService
	cache      *reconcileCache
}

func NewReconciliationService(att *AttendanceService, sch *ScheduleService) *ReconciliationService {
	return &ReconciliationService{
		attendance: att,
		schedule:   sch,
		cache:      newReconcileCache(),
	}
}

// ReconcileResult агрегированная сверка по отделениям и группам.
type ReconcileResult struct {
	Date        string                      `json:"date"`
	TotalPlanned int                        `json:"totalPlanned"`
	TotalPresent int                        `json:"totalPresent"`
	TotalAbsent  int                        `json:"totalAbsent"`
	ByDepartment []ReconcileDepartmentStats `json:"byDepartment"`
}

type ReconcileDepartmentStats struct {
	Department  string                   `json:"department"`
	Planned     int                      `json:"planned"`
	Present     int                      `json:"present"`
	Absent      int                      `json:"absent"`
	ByGroup     []ReconcileGroupStats    `json:"byGroup"`
}

type ReconcileGroupStats struct {
	Group      string `json:"group"`
	Planned    int    `json:"planned"`
	Present    int    `json:"present"`
	Absent     int    `json:"absent"`
	Discipline string `json:"discipline,omitempty"` // на паре — какая дисциплина
	Teacher    string `json:"teacher,omitempty"`   // преподаватель
}

// ReconcileDay выполняет сверку за конкретную дату (YYYY-MM-DD).
// Логика:
//  - schedule.json говорит, какие студенты должны быть (planned).
//  - attendance.json говорит, у кого есть пропуски (missed > 0) в этот день.
//  - absent = planned - present (present = те, у кого нет пропусков в этот день).
func (s *ReconciliationService) ReconcileDay(date string) (*ReconcileResult, error) {
	if s.attendance == nil || s.schedule == nil {
		return nil, fmt.Errorf("reconciliation service is not fully configured")
	}

	if cached := s.cache.Get(date); cached != nil {
		return cached, nil
	}

	planned, err := s.schedule.GetPlannedForDate(date)
	if err != nil {
		return nil, err
	}

	flat, err := s.attendance.LoadFlatForDate(date)
	if err != nil {
		return nil, err
	}

	// Строим множество отсутствующих: ключ = dept|group|student (для vedomost дата уже подставлена в flat)
	absentSet := make(map[string]struct{})
	for _, rec := range flat {
		if rec.Date != date {
			continue
		}
		if rec.Missed <= 0 {
			continue
		}
		key := rec.Department + "|" + rec.Group + "|" + rec.Student
		absentSet[key] = struct{}{}
	}

	// Агрегируем по planned.
	deptAgg := make(map[string]*ReconcileDepartmentStats)

	totalPlanned := 0
	totalPresent := 0

	for _, p := range planned {
		key := p.Department + "|" + p.Group + "|" + p.Student
		totalPlanned++

		if p.Department == "" {
			// Не выводим "Неизвестное отделение" в списке — учитываем только в общих итогах
			if _, absent := absentSet[key]; !absent {
				totalPresent++
			}
			continue
		}

		deptKey := p.Department
		dept := deptAgg[deptKey]
		if dept == nil {
			dept = &ReconcileDepartmentStats{
				Department: deptKey,
				ByGroup:    []ReconcileGroupStats{},
			}
			deptAgg[deptKey] = dept
		}

		var grp *ReconcileGroupStats
		for i := range dept.ByGroup {
			if dept.ByGroup[i].Group == p.Group {
				grp = &dept.ByGroup[i]
				break
			}
		}
		if grp == nil {
			dept.ByGroup = append(dept.ByGroup, ReconcileGroupStats{
				Group: p.Group,
			})
			grp = &dept.ByGroup[len(dept.ByGroup)-1]
		}

		grp.Planned++
		dept.Planned++

		if _, absent := absentSet[key]; absent {
			grp.Absent++
			dept.Absent++
			continue
		}

		grp.Present++
		dept.Present++
		totalPresent++
	}

	var byDept []ReconcileDepartmentStats
	for _, d := range deptAgg {
		byDept = append(byDept, *d)
	}

	result := &ReconcileResult{
		Date:        date,
		TotalPlanned: totalPlanned,
		TotalPresent: totalPresent,
		TotalAbsent:  totalPlanned - totalPresent,
		ByDepartment: byDept,
	}

	s.cache.Set(date, result)

	return result, nil
}

// ReconcileDayLesson — сверка за дату и номер пары (1–6). Для сценария «первая пара — кто где есть/нет».
func (s *ReconciliationService) ReconcileDayLesson(date string, lessonNumber int) (*ReconcileResult, error) {
	if s.attendance == nil || s.schedule == nil {
		return nil, fmt.Errorf("reconciliation service is not fully configured")
	}

	cacheKey := date + "_lesson_" + fmt.Sprintf("%d", lessonNumber)
	if cached := s.cache.Get(cacheKey); cached != nil {
		return cached, nil
	}

	planned, err := s.schedule.GetPlannedForDateAndLesson(date, lessonNumber)
	if err != nil {
		return nil, err
	}

	flat, err := s.attendance.LoadFlatForDate(date)
	if err != nil {
		return nil, err
	}

	absentSet := make(map[string]struct{})
	for _, rec := range flat {
		if rec.Date != date {
			continue
		}
		if rec.Missed <= 0 {
			continue
		}
		key := rec.Department + "|" + rec.Group + "|" + rec.Student
		absentSet[key] = struct{}{}
	}

	deptAgg := make(map[string]*ReconcileDepartmentStats)
	totalPlanned := 0
	totalPresent := 0

	for _, p := range planned {
		key := p.Department + "|" + p.Group + "|" + p.Student
		totalPlanned++

		if p.Department == "" {
			if _, absent := absentSet[key]; !absent {
				totalPresent++
			}
			continue
		}

		deptKey := p.Department
		dept := deptAgg[deptKey]
		if dept == nil {
			dept = &ReconcileDepartmentStats{
				Department: deptKey,
				ByGroup:    []ReconcileGroupStats{},
			}
			deptAgg[deptKey] = dept
		}

		var grp *ReconcileGroupStats
		for i := range dept.ByGroup {
			if dept.ByGroup[i].Group == p.Group {
				grp = &dept.ByGroup[i]
				break
			}
		}
		if grp == nil {
			grp = &ReconcileGroupStats{
				Group:      p.Group,
				Discipline: p.Discipline,
				Teacher:    p.Teacher,
			}
			dept.ByGroup = append(dept.ByGroup, *grp)
			grp = &dept.ByGroup[len(dept.ByGroup)-1]
		}

		grp.Planned++
		dept.Planned++

		if _, absent := absentSet[key]; absent {
			grp.Absent++
			dept.Absent++
			continue
		}
		grp.Present++
		dept.Present++
		totalPresent++
	}

	var byDept []ReconcileDepartmentStats
	for _, d := range deptAgg {
		byDept = append(byDept, *d)
	}

	result := &ReconcileResult{
		Date:         date,
		TotalPlanned: totalPlanned,
		TotalPresent: totalPresent,
		TotalAbsent:  totalPlanned - totalPresent,
		ByDepartment: byDept,
	}
	s.cache.Set(cacheKey, result)
	return result, nil
}

// LessonGroupStudent — студент на паре: есть/нет
type LessonGroupStudent struct {
	Student string `json:"student"`
	Present bool   `json:"present"`
}

// LessonGroupDetail — по одной группе на пару: дисциплина, преподаватель, список студентов
type LessonGroupDetail struct {
	Group      string               `json:"group"`
	Department string               `json:"department"`
	Discipline string               `json:"discipline"`
	Teacher    string               `json:"teacher"`
	Planned    int                  `json:"planned"`
	Present    int                  `json:"present"`
	Absent     int                  `json:"absent"`
	Students   []LessonGroupStudent `json:"students"`
}

// ReconcileDayLessonGroup — сверка по группе на конкретную пару: кто есть, кто нет.
func (s *ReconciliationService) ReconcileDayLessonGroup(date string, lessonNumber int, group string) (*LessonGroupDetail, error) {
	if s.attendance == nil || s.schedule == nil {
		return nil, fmt.Errorf("reconciliation service is not fully configured")
	}

	planned, err := s.schedule.GetPlannedForDateAndLesson(date, lessonNumber)
	if err != nil {
		return nil, err
	}

	flat, err := s.attendance.LoadFlatForDate(date)
	if err != nil {
		return nil, err
	}

	absentSet := make(map[string]struct{})
	for _, rec := range flat {
		if rec.Date != date {
			continue
		}
		if rec.Missed <= 0 {
			continue
		}
		key := rec.Department + "|" + rec.Group + "|" + rec.Student
		absentSet[key] = struct{}{}
	}

	var detail *LessonGroupDetail
	// Дедуплицируем студентов: ключ = Department|Group|Student
	seenStudents := make(map[string]bool)
	for _, p := range planned {
		if p.Group != group {
			continue
		}
		if detail == nil {
			detail = &LessonGroupDetail{
				Group:      p.Group,
				Department: p.Department,
				Discipline: p.Discipline,
				Teacher:    p.Teacher,
				Students:   []LessonGroupStudent{},
			}
		}
		studentKey := p.Department + "|" + p.Group + "|" + p.Student
		// Пропускаем дубликаты студентов
		if seenStudents[studentKey] {
			continue
		}
		seenStudents[studentKey] = true

		key := p.Department + "|" + p.Group + "|" + p.Student
		present := true
		if _, absent := absentSet[key]; absent {
			present = false
		}
		detail.Students = append(detail.Students, LessonGroupStudent{
			Student: p.Student,
			Present: present,
		})
		detail.Planned++
		if present {
			detail.Present++
		} else {
			detail.Absent++
		}
	}

	return detail, nil
}

// Простое in-memory кэширование результатов сверки по дате с TTL.

type reconcileCache struct {
	mu    sync.Mutex
	data  map[string]cacheEntry
	ttl   time.Duration
}

type cacheEntry struct {
	value     *ReconcileResult
	expiresAt time.Time
}

func newReconcileCache() *reconcileCache {
	return &reconcileCache{
		data: make(map[string]cacheEntry),
		ttl:  5 * time.Minute,
	}
}

func (c *reconcileCache) Get(date string) *ReconcileResult {
	c.mu.Lock()
	defer c.mu.Unlock()

	entry, ok := c.data[date]
	if !ok || time.Now().After(entry.expiresAt) {
		if ok {
			delete(c.data, date)
		}
		return nil
	}
	return entry.value
}

func (c *reconcileCache) Set(date string, value *ReconcileResult) {
	c.mu.Lock()
	defer c.mu.Unlock()

	c.data[date] = cacheEntry{
		value:     value,
		expiresAt: time.Now().Add(c.ttl),
	}
}



