package services

import (
	"database/sql"
	"fmt"
	"time"

	"dashboard/internal/data"
)

// DashboardStats общая статистика для главного экрана
type DashboardStats struct {
	TotalStudents    int `json:"totalStudents"`    // Всего студентов в колледже
	PresentNow       int `json:"presentNow"`       // Присутствующих в текущий момент
	AbsentNow        int `json:"absentNow"`        // Отсутствующих в текущий момент
	AttendancePercent float64 `json:"attendancePercent"` // Процент посещаемости сейчас
}

// TodayLesson информация о паре для главного экрана
type TodayLesson struct {
	ID              int       `json:"id"`              // ID занятия
	Group           string    `json:"group"`           // Название группы
	Department      string    `json:"department"`      // Отделение
	Discipline      string    `json:"discipline"`      // Дисциплина
	Teacher         string    `json:"teacher"`         // Преподаватель
	DateTime        time.Time `json:"dateTime"`       // Дата и время начала
	Planned         int       `json:"planned"`        // Плановое количество студентов
	Present         int       `json:"present"`        // Фактическое количество присутствующих
	Percent         float64   `json:"percent"`         // Процент посещаемости
	Status          string    `json:"status"`          // Статус: "future" / "current" / "past"
	Color           string    `json:"color"`          // Цвет индикации: "green" / "yellow" / "red" / "gray"
}

// CurrentLesson текущая пара (если есть)
type CurrentLesson struct {
	Lesson      *TodayLesson `json:"lesson"`      // Информация о паре
	IsActive    bool         `json:"isActive"`    // Идёт ли пара сейчас
	TimeRemaining string     `json:"timeRemaining"` // Оставшееся время (если идёт)
}

// DashboardService предоставляет бизнес-логику для главного экрана
type DashboardService struct {
	db            *sql.DB
	attendanceService *AttendanceService
	studentsPath  string
	jsonStore     *data.JSONStore
}

func NewDashboardService(db *sql.DB) *DashboardService {
	return &DashboardService{
		db:         db,
		jsonStore:  data.NewJSONStore(),
	}
}

// SetAttendanceService устанавливает сервис посещаемости для работы с JSON
func (s *DashboardService) SetAttendanceService(attendanceService *AttendanceService) {
	s.attendanceService = attendanceService
}

// SetStudentsPath устанавливает путь к файлу students.json
func (s *DashboardService) SetStudentsPath(studentsPath string) {
	s.studentsPath = studentsPath
}

// getThresholds получает пороги из БД (или возвращает дефолтные)
func (s *DashboardService) getThresholds(thresholdType string) (greenThreshold, yellowThreshold float64) {
	if s.db == nil {
		// Дефолтные значения, если БД не подключена
		return 90.0, 70.0
	}

	var green, yellow float64
	err := s.db.QueryRow(`
		SELECT green_threshold, yellow_threshold
		FROM thresholds
		WHERE type = $1
	`, thresholdType).Scan(&green, &yellow)

	if err == sql.ErrNoRows || err != nil {
		// Если пороги не найдены или ошибка - используем дефолтные
		return 90.0, 70.0
	}

	return green, yellow
}

// GetStats возвращает общую статистику для главного экрана
func (s *DashboardService) GetStats() (*DashboardStats, error) {
	// Пробуем использовать JSON файлы, если доступны
	if s.attendanceService != nil && s.studentsPath != "" {
		return s.getStatsFromJSON()
	}

	// Fallback на БД, если доступна
	if s.db != nil {
		return s.getStatsFromDB()
	}

	return nil, fmt.Errorf("нет доступа к данным (ни БД, ни JSON файлы)")
}

// getStatsFromJSON получает статистику из JSON файлов
func (s *DashboardService) getStatsFromJSON() (*DashboardStats, error) {
	// Общее количество студентов из students.json
	type studentsRoot struct {
		TotalStudents int `json:"totalStudents"`
	}

	var totalStudents int
	if s.studentsPath != "" && s.jsonStore != nil {
		root, err := data.LoadJSON[studentsRoot](s.jsonStore, s.studentsPath)
		if err == nil {
			totalStudents = root.TotalStudents
		}
	}

	// Получаем данные о посещаемости из attendanceService
	var presentNow, absentNow int
	var attendancePercent float64

	if s.attendanceService != nil {
		_, flat, err := s.attendanceService.LoadFromJSON()
		if err == nil {
			// Подсчитываем присутствующих и отсутствующих из всех записей
			for _, record := range flat {
				if record.Missed == 0 {
					presentNow++
				} else {
					absentNow++
				}
			}

			// Процент посещаемости
			totalOnLessons := presentNow + absentNow
			if totalOnLessons > 0 {
				attendancePercent = float64(presentNow) * 100.0 / float64(totalOnLessons)
			}
		}
	}

	return &DashboardStats{
		TotalStudents:    totalStudents,
		PresentNow:       presentNow,
		AbsentNow:        absentNow,
		AttendancePercent: attendancePercent,
	}, nil
}

// getStatsFromDB получает статистику из БД (fallback)
func (s *DashboardService) getStatsFromDB() (*DashboardStats, error) {
	// Общее количество студентов
	var totalStudents int
	err := s.db.QueryRow("SELECT COUNT(*) FROM students").Scan(&totalStudents)
	if err != nil {
		return nil, fmt.Errorf("ошибка получения общего количества студентов: %w", err)
	}

	// Присутствующих в текущий момент (на текущих парах)
	now := time.Now()
	var presentNow int
	err = s.db.QueryRow(`
		SELECT COUNT(DISTINCT la.student_id)
		FROM lessons l
		JOIN lesson_attendance la ON la.lesson_id = l.id
		WHERE la.attendance = true
		  AND l.date_time <= $1
		  AND l.date_time + INTERVAL '90 minutes' >= $1
	`, now).Scan(&presentNow)
	if err != nil && err != sql.ErrNoRows {
		return nil, fmt.Errorf("ошибка получения присутствующих: %w", err)
	}

	// Отсутствующих в текущий момент
	var absentNow int
	err = s.db.QueryRow(`
		SELECT COUNT(DISTINCT la.student_id)
		FROM lessons l
		JOIN lesson_attendance la ON la.lesson_id = l.id
		WHERE la.attendance = false
		  AND l.date_time <= $1
		  AND l.date_time + INTERVAL '90 minutes' >= $1
	`, now).Scan(&absentNow)
	if err != nil && err != sql.ErrNoRows {
		return nil, fmt.Errorf("ошибка получения отсутствующих: %w", err)
	}

	// Процент посещаемости
	var attendancePercent float64
	totalOnLessons := presentNow + absentNow
	if totalOnLessons > 0 {
		attendancePercent = float64(presentNow) * 100.0 / float64(totalOnLessons)
	}

	return &DashboardStats{
		TotalStudents:    totalStudents,
		PresentNow:       presentNow,
		AbsentNow:        absentNow,
		AttendancePercent: attendancePercent,
	}, nil
}

// GetTodayLessons возвращает список всех пар за сегодня с агрегатами
func (s *DashboardService) GetTodayLessons() ([]TodayLesson, error) {
	if s.db == nil {
		return nil, fmt.Errorf("БД не подключена")
	}

	today := time.Now().Format("2006-01-02")
	now := time.Now()

	rows, err := s.db.Query(`
		SELECT
			l.id,
			g.name AS group_name,
			d.name AS department_name,
			l.discipline,
			l.teacher,
			l.date_time,
			(SELECT COUNT(*) FROM students WHERE group_id = g.id) AS planned,
			COALESCE(SUM(CASE WHEN la.attendance THEN 1 ELSE 0 END), 0) AS present
		FROM lessons l
		JOIN groups g ON g.id = l.group_id
		JOIN departments d ON d.id = g.department_id
		LEFT JOIN lesson_attendance la ON la.lesson_id = l.id
		WHERE l.date_time::date = $1::date
		GROUP BY l.id, g.name, g.id, d.name, l.discipline, l.teacher, l.date_time
		ORDER BY l.date_time, g.name, l.discipline
	`, today)
	if err != nil {
		return nil, fmt.Errorf("ошибка выборки занятий: %w", err)
	}
	defer rows.Close()

	var result []TodayLesson

	for rows.Next() {
		var lesson TodayLesson
		var planned, present int

		if err := rows.Scan(
			&lesson.ID,
			&lesson.Group,
			&lesson.Department,
			&lesson.Discipline,
			&lesson.Teacher,
			&lesson.DateTime,
			&planned,
			&present,
		); err != nil {
			return nil, fmt.Errorf("ошибка чтения строки: %w", err)
		}

		lesson.Planned = planned
		lesson.Present = present
		if planned > 0 {
			lesson.Percent = float64(present) * 100.0 / float64(planned)
		}

		// Определяем статус пары
		lessonStart := lesson.DateTime
		lessonEnd := lessonStart.Add(90 * time.Minute) // Пара длится 90 минут

		if now.Before(lessonStart) {
			lesson.Status = "future"
		} else if now.After(lessonEnd) {
			lesson.Status = "past"
		} else {
			lesson.Status = "current"
		}

		// Определяем цвет индикации (читаем пороги из БД)
		greenThreshold, yellowThreshold := s.getThresholds("lesson")
		lesson.Color = s.getColorForPercent(lesson.Percent, lesson.Status, greenThreshold, yellowThreshold)

		result = append(result, lesson)
	}

	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("ошибка итерации по строкам: %w", err)
	}

	return result, nil
}

// GetCurrentLesson возвращает текущую пару (если есть)
func (s *DashboardService) GetCurrentLesson() (*CurrentLesson, error) {
	if s.db == nil {
		return nil, fmt.Errorf("БД не подключена")
	}

	now := time.Now()
	today := now.Format("2006-01-02")

	// Ищем пару, которая идёт сейчас
	var lesson TodayLesson
	var planned, present int
	var lessonID sql.NullInt64

	err := s.db.QueryRow(`
		SELECT
			l.id,
			g.name AS group_name,
			d.name AS department_name,
			l.discipline,
			l.teacher,
			l.date_time,
			(SELECT COUNT(*) FROM students WHERE group_id = g.id) AS planned,
			COALESCE(SUM(CASE WHEN la.attendance THEN 1 ELSE 0 END), 0) AS present
		FROM lessons l
		JOIN groups g ON g.id = l.group_id
		JOIN departments d ON d.id = g.department_id
		LEFT JOIN lesson_attendance la ON la.lesson_id = l.id
		WHERE l.date_time::date = $1::date
		  AND l.date_time <= $2
		  AND l.date_time + INTERVAL '90 minutes' >= $2
		GROUP BY l.id, g.name, g.id, d.name, l.discipline, l.teacher, l.date_time
		ORDER BY l.date_time DESC
		LIMIT 1
	`, today, now).Scan(
		&lessonID,
		&lesson.Group,
		&lesson.Department,
		&lesson.Discipline,
		&lesson.Teacher,
		&lesson.DateTime,
		&planned,
		&present,
	)

	if err == sql.ErrNoRows {
		// Нет текущей пары
		return &CurrentLesson{
			Lesson:   nil,
			IsActive: false,
		}, nil
	}
	if err != nil {
		return nil, fmt.Errorf("ошибка получения текущей пары: %w", err)
	}

	lesson.ID = int(lessonID.Int64)
	lesson.Planned = planned
	lesson.Present = present
	if planned > 0 {
		lesson.Percent = float64(present) * 100.0 / float64(planned)
	}
	lesson.Status = "current"
	greenThreshold, yellowThreshold := s.getThresholds("lesson")
	lesson.Color = s.getColorForPercent(lesson.Percent, lesson.Status, greenThreshold, yellowThreshold)

	// Вычисляем оставшееся время
	lessonEnd := lesson.DateTime.Add(90 * time.Minute)
	timeRemaining := ""
	if now.Before(lessonEnd) {
		remaining := lessonEnd.Sub(now)
		minutes := int(remaining.Minutes())
		timeRemaining = fmt.Sprintf("%d мин", minutes)
	}

	return &CurrentLesson{
		Lesson:        &lesson,
		IsActive:      true,
		TimeRemaining: timeRemaining,
	}, nil
}

// getColorForPercent определяет цвет индикации на основе процента посещаемости
func (s *DashboardService) getColorForPercent(percent float64, status string, greenThreshold, yellowThreshold float64) string {
	// Если пара ещё не началась или нет данных
	if status == "future" || percent == 0 {
		return "gray"
	}

	// Определяем цвет по порогам
	if percent >= greenThreshold {
		return "green"
	} else if percent >= yellowThreshold {
		return "yellow"
	} else {
		return "red"
	}
}
