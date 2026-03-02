package models

import "time"

// Department модель отделения
type Department struct {
	ID        int       `json:"id" db:"id"`
	Name      string    `json:"name" db:"name"`
	CreatedAt time.Time `json:"created_at" db:"created_at"`
	UpdatedAt time.Time `json:"updated_at" db:"updated_at"`
}

// Group модель группы
type Group struct {
	ID           int       `json:"id" db:"id"`
	DepartmentID int       `json:"department_id" db:"department_id"`
	Name         string    `json:"name" db:"name"`
	CreatedAt    time.Time `json:"created_at" db:"created_at"`
	UpdatedAt    time.Time `json:"updated_at" db:"updated_at"`
}

// Student модель студента
type Student struct {
	ID        int       `json:"id" db:"id"`
	GroupID   int       `json:"group_id" db:"group_id"`
	FullName  string    `json:"full_name" db:"full_name"`
	CreatedAt time.Time `json:"created_at" db:"created_at"`
	UpdatedAt time.Time `json:"updated_at" db:"updated_at"`
}

// Attendance модель посещаемости
type Attendance struct {
	ID          int       `json:"id" db:"id"`
	StudentID   int       `json:"student_id" db:"student_id"`
	Date        time.Time `json:"date" db:"date"`
	MissedHours int       `json:"missed_hours" db:"missed_hours"`
	CreatedAt   time.Time `json:"created_at" db:"created_at"`
	UpdatedAt   time.Time `json:"updated_at" db:"updated_at"`
}

// Lesson модель отдельного занятия (пара)
type Lesson struct {
	ID        int       `json:"id" db:"id"`
	GroupID   int       `json:"group_id" db:"group_id"`
	DateTime  time.Time `json:"date_time" db:"date_time"`
	Discipline string   `json:"discipline" db:"discipline"`
	Teacher   string    `json:"teacher" db:"teacher"`
	CreatedAt time.Time `json:"created_at" db:"created_at"`
	UpdatedAt time.Time `json:"updated_at" db:"updated_at"`
}

// LessonAttendance модель посещаемости конкретного занятия
type LessonAttendance struct {
	ID         int       `json:"id" db:"id"`
	LessonID   int       `json:"lesson_id" db:"lesson_id"`
	StudentID  int       `json:"student_id" db:"student_id"`
	Attendance bool      `json:"attendance" db:"attendance"`
	CreatedAt  time.Time `json:"created_at" db:"created_at"`
	UpdatedAt  time.Time `json:"updated_at" db:"updated_at"`
}

// Specialty модель специальности (для summary)
type Specialty struct {
	ID           int       `json:"id" db:"id"`
	DepartmentID int       `json:"department_id" db:"department_id"`
	Name         string    `json:"name" db:"name"`
	TotalMissed  int       `json:"total_missed" db:"total_missed"`
	CreatedAt    time.Time `json:"created_at" db:"created_at"`
	UpdatedAt    time.Time `json:"updated_at" db:"updated_at"`
}

// SummaryGroup модель группы в ведомости
type SummaryGroup struct {
	ID          int       `json:"id" db:"id"`
	SpecialtyID int       `json:"specialty_id" db:"specialty_id"`
	Name        string    `json:"name" db:"name"`
	TotalMissed int       `json:"total_missed" db:"total_missed"`
	CreatedAt   time.Time `json:"created_at" db:"created_at"`
	UpdatedAt   time.Time `json:"updated_at" db:"updated_at"`
}

// SummaryStudent модель студента в ведомости
type SummaryStudent struct {
	ID            int       `json:"id" db:"id"`
	SummaryGroupID int      `json:"summary_group_id" db:"summary_group_id"`
	FullName      string    `json:"full_name" db:"full_name"`
	MissedTotal   int       `json:"missed_total" db:"missed_total"`
	MissedBad     int       `json:"missed_bad" db:"missed_bad"`
	MissedExcused int       `json:"missed_excused" db:"missed_excused"`
	CreatedAt     time.Time `json:"created_at" db:"created_at"`
	UpdatedAt     time.Time `json:"updated_at" db:"updated_at"`
}

// VedomostRoot — корень vedomost.json (period из шапки + массив отделений). Обратная совместимость: если файл — массив, читаем как []VedomostDepartment.
type VedomostRoot struct {
	Period      string               `json:"period"`
	Departments []VedomostDepartment `json:"departments"`
}

// VedomostDepartment — отделение в vedomost.json (специальности → группы → студенты с missedTotal).
type VedomostDepartment struct {
	Department  string              `json:"department"`
	TotalMissed int                 `json:"totalMissed"`
	Specialties []VedomostSpecialty `json:"specialties"`
}
type VedomostSpecialty struct {
	Specialty   string           `json:"specialty"`
	TotalMissed int              `json:"totalMissed"`
	Groups      []VedomostGroup  `json:"groups"`
}
type VedomostGroup struct {
	Group       string             `json:"group"`
	TotalMissed int                `json:"totalMissed"`
	Students    []VedomostStudent  `json:"students"`
}
type VedomostStudent struct {
	Student       string `json:"student"`
	MissedTotal   int    `json:"missedTotal"`
	MissedBad     int    `json:"missedBad"`
	MissedExcused int    `json:"missedExcused"`
}

// JSON модели для работы с JSON файлами (attendance.json, schedule.json, vedomost.json)

// AttendanceRecordJSON запись посещаемости в JSON
type AttendanceRecordJSON struct {
	Date   string `json:"date"`
	Missed int    `json:"missed"`
}

// StudentJSON студент в JSON
type StudentJSON struct {
	StudentID  string                `json:"studentId,omitempty"`
	Student    string                `json:"student"`
	Attendance []AttendanceRecordJSON `json:"attendance"`
}

// GroupJSON группа в JSON
type GroupJSON struct {
	GroupID  string        `json:"groupId,omitempty"`
	Group    string        `json:"group"`
	Students []StudentJSON `json:"students"`
}

// DepartmentJSON отделение в JSON
type DepartmentJSON struct {
	DepartmentID string      `json:"departmentId,omitempty"`
	Department   string      `json:"department"`
	Groups       []GroupJSON `json:"groups"`
}

// ScheduleJSON описывает структуру файла schedule.json
type ScheduleJSON struct {
	Period        string              `json:"period"`
	Groups        []ScheduleGroupJSON `json:"groups"`
	TotalGroups   int                 `json:"totalGroups"`
	TotalStudents int                 `json:"totalStudents"`
}

type ScheduleGroupJSON struct {
	Group         string                `json:"group"`
	Department    string                `json:"department"`
	Students      []ScheduleStudentJSON `json:"students"`
	TotalStudents int                   `json:"totalStudents"`
}

type ScheduleStudentJSON struct {
	StudentName   string                 `json:"studentName"`
	NumberInGroup int                    `json:"numberInGroup"`
	Records       []ScheduleLessonRecord `json:"records"`
	TotalCount    int                    `json:"totalCount"`
}

type ScheduleLessonRecord struct {
	Date         string `json:"date"`
	LessonNumber int    `json:"lessonNumber,omitempty"` // Номер пары (1-6), для фильтра "первая пара"
	Discipline   string `json:"discipline"`
	Teacher      string `json:"teacher"`
	Attendance   bool   `json:"attendance"`
}

// Schedule модель расписания (связь пар с группами по дням недели)
type Schedule struct {
	ID          int       `json:"id" db:"id"`
	GroupID     int       `json:"group_id" db:"group_id"`
	DayOfWeek   int       `json:"day_of_week" db:"day_of_week"`   // 0 = воскресенье, 1 = понедельник, ..., 6 = суббота
	LessonNumber int      `json:"lesson_number" db:"lesson_number"` // Номер пары (1-8)
	Discipline  string    `json:"discipline" db:"discipline"`
	Teacher     string    `json:"teacher" db:"teacher"`
	StartTime   string    `json:"start_time" db:"start_time"`   // Время начала (HH:MM:SS)
	EndTime     string    `json:"end_time" db:"end_time"`       // Время окончания (HH:MM:SS)
	CreatedAt   time.Time `json:"created_at" db:"created_at"`
	UpdatedAt   time.Time `json:"updated_at" db:"updated_at"`
}

// Threshold модель порогов цветовой индикации
type Threshold struct {
	ID             int       `json:"id" db:"id"`
	Type           string    `json:"type" db:"type"`             // 'lesson', 'group', 'department'
	GreenThreshold float64   `json:"green_threshold" db:"green_threshold"`  // Верхний порог (>=)
	YellowThreshold float64  `json:"yellow_threshold" db:"yellow_threshold"` // Нижний порог (>=)
	CreatedAt      time.Time `json:"created_at" db:"created_at"`
	UpdatedAt      time.Time `json:"updated_at" db:"updated_at"`
}
