package services

import (
	"fmt"
	"net/http"
	"strconv"
	"strings"
	"time"

	"dashboard/internal/data"
	"dashboard/internal/models"
)

type AttendanceService struct {
	attendancePath string
	vedomostPath   string // приоритетный источник для сверки (reconcile); если задан — используем vedomost.json
	studentsPath   string
	jsonStore      *data.JSONStore
}

// NewAttendanceService: attendancePath — attendance.json; vedomostPath — vedomost.json для сверки (если не пусто, reconcile использует его вместо attendance).
func NewAttendanceService(attendancePath, studentsPath, vedomostPath string) *AttendanceService {
	return &AttendanceService{
		attendancePath: attendancePath,
		vedomostPath:   vedomostPath,
		studentsPath:   studentsPath,
		jsonStore:      data.NewJSONStore(),
	}
}

func (s *AttendanceService) loadContingent() (byDept map[string]int, byGroup map[string]map[string]int) {
	byDept = make(map[string]int)
	byGroup = make(map[string]map[string]int)
	if s.studentsPath == "" || s.jsonStore == nil {
		return byDept, byGroup
	}

	type studentsRoot struct {
		Departments []struct {
			Department string `json:"department"`
			Groups     []struct {
				Group    string        `json:"group"`
				Students []interface{} `json:"students"`
			} `json:"groups"`
		} `json:"departments"`
	}

	root, err := data.LoadJSON[studentsRoot](s.jsonStore, s.studentsPath)
	if err != nil {
		return byDept, byGroup
	}
	for _, d := range root.Departments {
		n := 0
		if byGroup[d.Department] == nil {
			byGroup[d.Department] = make(map[string]int)
		}
		for _, g := range d.Groups {
			c := len(g.Students)
			n += c
			byGroup[d.Department][g.Group] = c
		}
		if n > 0 {
			byDept[d.Department] = n
		}
	}
	return byDept, byGroup
}

func (s *AttendanceService) loadStudentNamesForGroup(department, group string) []string {
	if s.studentsPath == "" || s.jsonStore == nil {
		return nil
	}

	type studentsRoot struct {
		Departments []struct {
			Department string `json:"department"`
			Groups     []struct {
				Group    string `json:"group"`
				Students []struct {
					FullName string `json:"fullName"`
				} `json:"students"`
			} `json:"groups"`
		} `json:"departments"`
	}

	root, err := data.LoadJSON[studentsRoot](s.jsonStore, s.studentsPath)
	if err != nil {
		return nil
	}
	for _, d := range root.Departments {
		if d.Department != department {
			continue
		}
		for _, g := range d.Groups {
			if g.Group != group {
				continue
			}
			names := make([]string, 0, len(g.Students))
			for _, st := range g.Students {
				if st.FullName != "" {
					names = append(names, st.FullName)
				}
			}
			return names
		}
	}
	return nil
}

func (s *AttendanceService) LoadFromJSON() ([]models.DepartmentJSON, []models.FlatRecord, error) {
	if s.attendancePath == "" || s.jsonStore == nil {
		return nil, nil, fmt.Errorf("attendance service is not configured with paths")
	}

	departments, err := data.LoadJSON[[]models.DepartmentJSON](s.jsonStore, s.attendancePath)
	if err != nil {
		return nil, nil, err
	}

	flat := models.Flatten(departments)
	return departments, flat, nil
}

// loadVedomostDepartments читает vedomost.json: поддерживает формат с period+departments и старый формат (только массив).
func loadVedomostDepartments(store *data.JSONStore, path string) ([]models.VedomostDepartment, error) {
	root, err := data.LoadJSON[models.VedomostRoot](store, path)
	if err == nil && len(root.Departments) > 0 {
		return root.Departments, nil
	}
	// Обратная совместимость: файл — массив отделений
	arr, err := data.LoadJSON[[]models.VedomostDepartment](store, path)
	if err != nil {
		return nil, err
	}
	return arr, nil
}

// LoadFlatForDate возвращает плоский список записей для указанной даты. Для сверки (reconcile) и истории:
// приоритет — attendance.json по выбранной дате (корректно по дням 24–28 и т.д.); если за эту дату записей нет — vedomost.
func (s *AttendanceService) LoadFlatForDate(date string) ([]models.FlatRecord, error) {
	if s.attendancePath != "" && s.jsonStore != nil {
		_, flat, err := s.LoadFromJSON()
		if err == nil {
			out := make([]models.FlatRecord, 0, len(flat)/7)
			for _, rec := range flat {
				if rec.Date == date {
					out = append(out, rec)
				}
			}
			if len(out) > 0 {
				return out, nil
			}
		}
	}
	if s.vedomostPath != "" && s.jsonStore != nil {
		departments, err := loadVedomostDepartments(s.jsonStore, s.vedomostPath)
		if err != nil {
			return nil, fmt.Errorf("загрузка vedomost: %w", err)
		}
		return models.VedomostToFlat(departments, date), nil
	}
	return []models.FlatRecord{}, nil
}

type FilterParams struct {
	Department string
	Group      string
	Student    string
	Date       string
	DateFrom   string
	DateTo     string
	Period     string
	Search     string
	MissedMin  int
}

func (s *AttendanceService) Filter(records []models.FlatRecord, params FilterParams) []models.FlatRecord {
	today := time.Now().Format("2006-01-02")
	var from, to string

	switch params.Period {
	case "7d":
		from = todayAdd(-7)
		to = today
	case "30d":
		from = todayAdd(-30)
		to = today
	case "90d":
		from = todayAdd(-90)
		to = today
	default:
		from = params.DateFrom
		to = params.DateTo
	}

	searchLower := ""
	if params.Search != "" {
		searchLower = strings.ToLower(params.Search)
	}

	out := make([]models.FlatRecord, 0, len(records))
	for _, rec := range records {
		if params.Department != "" && rec.Department != params.Department {
			continue
		}

		if params.Group != "" && rec.Group != params.Group {
			continue
		}

		if params.Student != "" && rec.Student != params.Student {
			continue
		}

		if searchLower != "" {
			ok := strings.Contains(strings.ToLower(rec.Department), searchLower) ||
				strings.Contains(strings.ToLower(rec.Group), searchLower) ||
				strings.Contains(strings.ToLower(rec.Student), searchLower)
			if !ok {
				continue
			}
		}

		if params.MissedMin >= 0 && rec.Missed < params.MissedMin {
			continue
		}

		if from != "" || to != "" {
			if from != "" && rec.Date < from {
				continue
			}
			if to != "" && rec.Date > to {
				continue
			}
		} else {
			if params.Date == "today" {
				if rec.Date != today {
					continue
				}
			} else if params.Date != "" && rec.Date != params.Date {
				continue
			}
		}

		out = append(out, rec)
	}

	return out
}

func ParseFilterParams(r *http.Request) FilterParams {
	q := r.URL.Query()
	missedMin := -1
	if s := q.Get("missed_min"); s != "" {
		if n, err := strconv.Atoi(s); err == nil && n >= 0 {
			missedMin = n
		}
	}

	return FilterParams{
		Department: strings.TrimSpace(q.Get("department")),
		Group:      strings.TrimSpace(q.Get("group")),
		Student:    strings.TrimSpace(q.Get("student")),
		Date:       strings.TrimSpace(q.Get("date")),
		DateFrom:   strings.TrimSpace(q.Get("date_from")),
		DateTo:     strings.TrimSpace(q.Get("date_to")),
		Period:     strings.TrimSpace(q.Get("period")),
		Search:     strings.TrimSpace(q.Get("q")),
		MissedMin:  missedMin,
	}
}

func todayAdd(days int) string {
	t := time.Now().AddDate(0, 0, days)
	return t.Format("2006-01-02")
}

type SummaryResponse struct {
	TotalStudents int            `json:"total_students"`
	Present       int            `json:"present"`
	Absent        int            `json:"absent"`
	ByDepartment  []DeptDrillItem `json:"by_department,omitempty"`
}

type DeptDrillItem struct {
	Department  string `json:"department"`
	Total       int    `json:"total"`
	Absent      int    `json:"absent"`
	MissedTotal int    `json:"missed_total"`
}

type GroupDrillItem struct {
	Group       string `json:"group"`
	Total       int    `json:"total"`
	Absent      int    `json:"absent"`
	MissedTotal int    `json:"missed_total"`
}

type StudentDrillItem struct {
	Student     string   `json:"student"`
	MissedTotal int      `json:"missed_total"`
	Records     int      `json:"records"`
	Dates       []string `json:"dates,omitempty"`
}

func (s *AttendanceService) BuildSummary(departments []models.DepartmentJSON, filtered []models.FlatRecord) SummaryResponse {
	byDept := totalByDept(departments)
	contingentDept, _ := s.loadContingent()
	absentSet := make(map[string]struct{})
	deptAbsent := make(map[string]int)
	deptMissed := make(map[string]int)

	for _, rec := range filtered {
		k := rec.Department + "\x00" + rec.Group + "\x00" + rec.Student
		if _, ok := absentSet[k]; !ok {
			absentSet[k] = struct{}{}
			deptAbsent[rec.Department]++
		}
		deptMissed[rec.Department] += rec.Missed
	}

	iter := byDept
	if len(contingentDept) > 0 {
		iter = contingentDept
	}
	var total int
	for _, n := range iter {
		total += n
	}

	absent := len(absentSet)
	present := total - absent
	if present < 0 {
		present = 0
	}

	var byDepartment []DeptDrillItem
	for dept, tot := range iter {
		abs := deptAbsent[dept]
		missed := deptMissed[dept]
		byDepartment = append(byDepartment, DeptDrillItem{
			Department:  dept,
			Total:       tot,
			Absent:      abs,
			MissedTotal: missed,
		})
	}

	return SummaryResponse{
		TotalStudents: total,
		Present:       present,
		Absent:        absent,
		ByDepartment:  byDepartment,
	}
}

func (s *AttendanceService) BuildDrillDepartments(departments []models.DepartmentJSON, filtered []models.FlatRecord) []DeptDrillItem {
	byDept := totalByDept(departments)
	contingentDept, _ := s.loadContingent()
	deptAbsent := make(map[string]int)
	deptMissed := make(map[string]int)
	seen := make(map[string]map[string]struct{})
	deptsInScope := make(map[string]struct{})

	for _, rec := range filtered {
		deptsInScope[rec.Department] = struct{}{}
		if seen[rec.Department] == nil {
			seen[rec.Department] = make(map[string]struct{})
		}
		k := rec.Group + "\x00" + rec.Student
		if _, ok := seen[rec.Department][k]; !ok {
			seen[rec.Department][k] = struct{}{}
			deptAbsent[rec.Department]++
		}
		deptMissed[rec.Department] += rec.Missed
	}

	iter := byDept
	if len(deptsInScope) > 0 {
		iter = make(map[string]int)
		for d := range deptsInScope {
			iter[d] = byDept[d]
		}
	}

	var out []DeptDrillItem
	for dept, tot := range iter {
		if contingentDept[dept] > 0 {
			tot = contingentDept[dept]
		}
		out = append(out, DeptDrillItem{
			Department:  dept,
			Total:       tot,
			Absent:      deptAbsent[dept],
			MissedTotal: deptMissed[dept],
		})
	}
	return out
}

func (s *AttendanceService) BuildDrillGroups(departments []models.DepartmentJSON, filtered []models.FlatRecord, department string) []GroupDrillItem {
	byGroup := totalByGroup(departments)
	_, contingentGroup := s.loadContingent()
	if byGroup[department] == nil && (contingentGroup[department] == nil || len(contingentGroup[department]) == 0) {
		return []GroupDrillItem{}
	}

	grpAbsent := make(map[string]int)
	grpMissed := make(map[string]int)
	seen := make(map[string]map[string]struct{})

	for _, rec := range filtered {
		if rec.Department != department {
			continue
		}
		if seen[rec.Group] == nil {
			seen[rec.Group] = make(map[string]struct{})
		}
		if _, ok := seen[rec.Group][rec.Student]; !ok {
			seen[rec.Group][rec.Student] = struct{}{}
			grpAbsent[rec.Group]++
		}
		grpMissed[rec.Group] += rec.Missed
	}

	iter := byGroup[department]
	if iter == nil {
		iter = contingentGroup[department]
	}
	if iter == nil {
		return []GroupDrillItem{}
	}
	var out []GroupDrillItem
	for grp, tot := range iter {
		if contingentGroup[department] != nil && contingentGroup[department][grp] > 0 {
			tot = contingentGroup[department][grp]
		}
		out = append(out, GroupDrillItem{
			Group:       grp,
			Total:       tot,
			Absent:      grpAbsent[grp],
			MissedTotal: grpMissed[grp],
		})
	}
	return out
}

func (s *AttendanceService) BuildDrillStudents(filtered []models.FlatRecord, department, group string) []StudentDrillItem {
	type agg struct {
		missed int
		dates  []string
	}
	m := make(map[string]*agg)
	for _, rec := range filtered {
		if rec.Department != department || rec.Group != group {
			continue
		}
		if m[rec.Student] == nil {
			m[rec.Student] = &agg{dates: []string{}}
		}
		m[rec.Student].missed += rec.Missed
		m[rec.Student].dates = append(m[rec.Student].dates, rec.Date)
	}

	names := s.loadStudentNamesForGroup(department, group)
	if len(names) == 0 {
		out := make([]StudentDrillItem, 0, len(m))
		for name, a := range m {
			out = append(out, StudentDrillItem{
				Student:     name,
				MissedTotal: a.missed,
				Records:     len(a.dates),
				Dates:       a.dates,
			})
		}
		return out
	}

	out := make([]StudentDrillItem, 0, len(names))
	for _, name := range names {
		item := StudentDrillItem{Student: name, MissedTotal: 0, Records: 0, Dates: []string{}}
		if a := m[name]; a != nil {
			item.MissedTotal = a.missed
			item.Records = len(a.dates)
			item.Dates = a.dates
		}
		out = append(out, item)
	}
	return out
}

func totalByDept(departments []models.DepartmentJSON) map[string]int {
	m := make(map[string]int)
	for _, d := range departments {
		n := 0
		for _, g := range d.Groups {
			n += len(g.Students)
		}
		if n > 0 {
			m[d.Department] = n
		}
	}
	return m
}

func totalByGroup(departments []models.DepartmentJSON) map[string]map[string]int {
	m := make(map[string]map[string]int)
	for _, d := range departments {
		if m[d.Department] == nil {
			m[d.Department] = make(map[string]int)
		}
		for _, g := range d.Groups {
			m[d.Department][g.Group] = len(g.Students)
		}
	}
	return m
}
