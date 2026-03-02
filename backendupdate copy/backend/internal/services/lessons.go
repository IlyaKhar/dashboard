package services

import (
	"database/sql"
	"fmt"
	"time"
)

// LessonSummary агрегированная информация по занятию для дашборда
type LessonSummary struct {
	Group      string    `json:"group"`
	Department string    `json:"department"`
	Discipline string    `json:"discipline"`
	DateTime   time.Time `json:"dateTime"`
	Planned    int       `json:"planned"`
	Present    int       `json:"present"`
	Percent    float64   `json:"percent"`
}

// LessonsService предоставляет бизнес-логику для работы с lessons/* таблицами
type LessonsService struct {
	db *sql.DB
}

func NewLessonsService(db *sql.DB) *LessonsService {
	return &LessonsService{db: db}
}

// GetDayLessons возвращает агрегаты по занятиям за день
func (s *LessonsService) GetDayLessons(date time.Time) ([]LessonSummary, error) {
	if s.db == nil {
		return nil, fmt.Errorf("БД не подключена")
	}

	day := date.Format("2006-01-02")

	rows, err := s.db.Query(`
		SELECT
			g.name AS group_name,
			d.name AS department_name,
			l.discipline,
			l.date_time,
			COUNT(DISTINCT st.id) AS planned,
			COALESCE(SUM(CASE WHEN la.attendance THEN 1 ELSE 0 END), 0) AS present
		FROM lessons l
		JOIN groups g ON g.id = l.group_id
		JOIN departments d ON d.id = g.department_id
		LEFT JOIN lesson_attendance la ON la.lesson_id = l.id
		LEFT JOIN students st ON st.id = la.student_id
		WHERE l.date_time::date = $1::date
		GROUP BY g.name, d.name, l.discipline, l.date_time
		ORDER BY l.date_time, g.name, l.discipline
	`, day)
	if err != nil {
		return nil, fmt.Errorf("ошибка выборки занятий: %w", err)
	}
	defer rows.Close()

	var result []LessonSummary

	for rows.Next() {
		var sRow LessonSummary
		var planned, present int

		if err := rows.Scan(
			&sRow.Group,
			&sRow.Department,
			&sRow.Discipline,
			&sRow.DateTime,
			&planned,
			&present,
		); err != nil {
			return nil, fmt.Errorf("ошибка чтения строки: %w", err)
		}

		sRow.Planned = planned
		sRow.Present = present
		if planned > 0 {
			sRow.Percent = float64(present) * 100.0 / float64(planned)
		}

		result = append(result, sRow)
	}

	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("ошибка итерации по строкам: %w", err)
	}

	return result, nil
}

