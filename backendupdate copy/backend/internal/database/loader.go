package database

import (
	"database/sql"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"time"
)

// Loader загружает JSON данные в БД
type Loader struct {
	db *sql.DB
}

func NewLoader(db *sql.DB) *Loader {
	return &Loader{db: db}
}

// LoadAttendance загружает данные посещаемости из JSON в БД
func (l *Loader) LoadAttendance(jsonPath string) error {
	if l.db == nil {
		return fmt.Errorf("БД не подключена")
	}

	log.Printf("[Database] Загрузка посещаемости из %s...", jsonPath)

	// Читаем JSON файл
	data, err := os.ReadFile(jsonPath)
	if err != nil {
		return fmt.Errorf("ошибка чтения файла: %v", err)
	}

	// Парсим JSON
	var departments []struct {
		Department string `json:"department"`
		Groups     []struct {
			Group    string `json:"group"`
			Students []struct {
				Student    string `json:"student"`
				Attendance []struct {
					Date   string `json:"date"`
					Missed int    `json:"missed"`
				} `json:"attendance"`
			} `json:"students"`
		} `json:"groups"`
	}

	if err := json.Unmarshal(data, &departments); err != nil {
		return fmt.Errorf("ошибка парсинга JSON: %v", err)
	}

	// Начинаем транзакцию
	tx, err := l.db.Begin()
	if err != nil {
		return fmt.Errorf("ошибка начала транзакции: %v", err)
	}
	defer tx.Rollback()

	// Очищаем старые данные (опционально - можно закомментировать для инкрементального обновления)
	// if _, err := tx.Exec("TRUNCATE TABLE attendance, students, groups, departments CASCADE"); err != nil {
	// 	return fmt.Errorf("ошибка очистки данных: %v", err)
	// }

	// Загружаем данные
	for _, dept := range departments {
		// Вставляем или получаем отделение
		var deptID int
		err := tx.QueryRow(
			`INSERT INTO departments (name) VALUES ($1) 
			 ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name 
			 RETURNING id`,
			dept.Department,
		).Scan(&deptID)
		if err != nil {
			return fmt.Errorf("ошибка вставки отделения %s: %v", dept.Department, err)
		}

		for _, group := range dept.Groups {
			// Вставляем или получаем группу
			var groupID int
			err := tx.QueryRow(
				`INSERT INTO groups (department_id, name) VALUES ($1, $2) 
				 ON CONFLICT (department_id, name) DO UPDATE SET name = EXCLUDED.name 
				 RETURNING id`,
				deptID, group.Group,
			).Scan(&groupID)
			if err != nil {
				return fmt.Errorf("ошибка вставки группы %s: %v", group.Group, err)
			}

			for _, student := range group.Students {
				// Вставляем или получаем студента
				var studentID int
				err := tx.QueryRow(
					`INSERT INTO students (group_id, full_name) VALUES ($1, $2) 
					 ON CONFLICT (group_id, full_name) DO UPDATE SET full_name = EXCLUDED.full_name 
					 RETURNING id`,
					groupID, student.Student,
				).Scan(&studentID)
				if err != nil {
					return fmt.Errorf("ошибка вставки студента %s: %v", student.Student, err)
				}

				// Вставляем записи посещаемости
				for _, att := range student.Attendance {
					// Парсим дату
					date, err := time.Parse("2006-01-02", att.Date)
					if err != nil {
						log.Printf("[Database] Предупреждение: неверный формат даты %s: %v", att.Date, err)
						continue
					}

					_, err = tx.Exec(
						`INSERT INTO attendance (student_id, date, missed_hours) 
						 VALUES ($1, $2, $3)
						 ON CONFLICT (student_id, date) 
						 DO UPDATE SET missed_hours = EXCLUDED.missed_hours`,
						studentID, date, att.Missed,
					)
					if err != nil {
						return fmt.Errorf("ошибка вставки посещаемости: %v", err)
					}
				}
			}
		}
	}

	// Коммитим транзакцию
	if err := tx.Commit(); err != nil {
		return fmt.Errorf("ошибка коммита транзакции: %v", err)
	}

	log.Printf("[Database] Посещаемость загружена. Отделений: %d", len(departments))
	return nil
}

// LoadStatement загружает данные ведомости из JSON в БД
func (l *Loader) LoadStatement(jsonPath string) error {
	if l.db == nil {
		return fmt.Errorf("БД не подключена")
	}

	log.Printf("[Database] Загрузка ведомости из %s...", jsonPath)

	// Читаем JSON файл
	data, err := os.ReadFile(jsonPath)
	if err != nil {
		return fmt.Errorf("ошибка чтения файла: %v", err)
	}

	// Парсим JSON
	var departments []struct {
		Department  string `json:"department"`
		TotalMissed int    `json:"totalMissed"`
		Specialties []struct {
			Specialty   string `json:"specialty"`
			TotalMissed int    `json:"totalMissed"`
			Groups      []struct {
				Group       string `json:"group"`
				TotalMissed int    `json:"totalMissed"`
				Students    []struct {
					Student       string `json:"student"`
					MissedTotal   int    `json:"missedTotal"`
					MissedBad     int    `json:"missedBad"`
					MissedExcused int    `json:"missedExcused"`
				} `json:"students"`
			} `json:"groups"`
		} `json:"specialties"`
	}

	if err := json.Unmarshal(data, &departments); err != nil {
		return fmt.Errorf("ошибка парсинга JSON: %v", err)
	}

	// Начинаем транзакцию
	tx, err := l.db.Begin()
	if err != nil {
		return fmt.Errorf("ошибка начала транзакции: %v", err)
	}
	defer tx.Rollback()

	// Очищаем старые данные summary (опционально)
	// if _, err := tx.Exec("TRUNCATE TABLE summary_students, summary_groups, specialties CASCADE"); err != nil {
	// 	return fmt.Errorf("ошибка очистки summary данных: %v", err)
	// }

	// Загружаем данные
	for _, dept := range departments {
		// Получаем ID отделения (должно существовать из attendance)
		var deptID int
		err := tx.QueryRow("SELECT id FROM departments WHERE name = $1", dept.Department).Scan(&deptID)
		if err != nil {
			// Если отделение не найдено, создаём его
			err = tx.QueryRow(
				`INSERT INTO departments (name) VALUES ($1) 
				 ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name 
				 RETURNING id`,
				dept.Department,
			).Scan(&deptID)
			if err != nil {
				return fmt.Errorf("ошибка получения/создания отделения %s: %v", dept.Department, err)
			}
		}

		for _, spec := range dept.Specialties {
			// Вставляем или получаем специальность
			var specID int
			err := tx.QueryRow(
				`INSERT INTO specialties (department_id, name, total_missed) VALUES ($1, $2, $3) 
				 ON CONFLICT (department_id, name) 
				 DO UPDATE SET total_missed = EXCLUDED.total_missed 
				 RETURNING id`,
				deptID, spec.Specialty, spec.TotalMissed,
			).Scan(&specID)
			if err != nil {
				return fmt.Errorf("ошибка вставки специальности %s: %v", spec.Specialty, err)
			}

			for _, group := range spec.Groups {
				// Вставляем или получаем группу summary
				var summaryGroupID int
				err := tx.QueryRow(
					`INSERT INTO summary_groups (specialty_id, name, total_missed) VALUES ($1, $2, $3) 
					 ON CONFLICT (specialty_id, name) 
					 DO UPDATE SET total_missed = EXCLUDED.total_missed 
					 RETURNING id`,
					specID, group.Group, group.TotalMissed,
				).Scan(&summaryGroupID)
				if err != nil {
					return fmt.Errorf("ошибка вставки summary группы %s: %v", group.Group, err)
				}

				for _, student := range group.Students {
					// Вставляем студента summary
					_, err = tx.Exec(
						`INSERT INTO summary_students (summary_group_id, full_name, missed_total, missed_bad, missed_excused) 
						 VALUES ($1, $2, $3, $4, $5)
						 ON CONFLICT (summary_group_id, full_name) 
						 DO UPDATE SET 
						 	missed_total = EXCLUDED.missed_total,
						 	missed_bad = EXCLUDED.missed_bad,
						 	missed_excused = EXCLUDED.missed_excused`,
						summaryGroupID, student.Student, student.MissedTotal, student.MissedBad, student.MissedExcused,
					)
					if err != nil {
						return fmt.Errorf("ошибка вставки summary студента %s: %v", student.Student, err)
					}
				}
			}
		}
	}

	// Коммитим транзакцию
	if err := tx.Commit(); err != nil {
		return fmt.Errorf("ошибка коммита транзакции: %v", err)
	}

	log.Printf("[Database] Ведомость загружена. Отделений: %d", len(departments))
	return nil
}

// LoadLessons загружает данные расписания и явки по занятиям из JSON в БД
func (l *Loader) LoadLessons(jsonPath string) error {
	if l.db == nil {
		return fmt.Errorf("БД не подключена")
	}

	log.Printf("[Database] Загрузка расписания занятий из %s...", jsonPath)

	data, err := os.ReadFile(jsonPath)
	if err != nil {
		return fmt.Errorf("ошибка чтения файла: %v", err)
	}

	// Структура соответствует LessonsOutput из converter/lessons.go
	var lessonsData struct {
		Period        string `json:"period"`
		Groups        []struct {
			Group         string `json:"group"`
			Department    string `json:"department"`
			TotalStudents int    `json:"totalStudents"`
			Students      []struct {
				StudentName   string `json:"studentName"`
				NumberInGroup int    `json:"numberInGroup"`
				Records       []struct {
					Date       string `json:"date"`
					Discipline string `json:"discipline"`
					Teacher    string `json:"teacher"`
					Attendance bool   `json:"attendance"`
				} `json:"records"`
			} `json:"students"`
		} `json:"groups"`
		TotalGroups   int `json:"totalGroups"`
		TotalStudents int `json:"totalStudents"`
	}

	if err := json.Unmarshal(data, &lessonsData); err != nil {
		return fmt.Errorf("ошибка парсинга JSON lessons: %v", err)
	}

	tx, err := l.db.Begin()
	if err != nil {
		return fmt.Errorf("ошибка начала транзакции: %v", err)
	}
	defer tx.Rollback()

	for _, g := range lessonsData.Groups {
		if g.Group == "" {
			continue
		}

		// Отделение
		deptName := g.Department
		if deptName == "" {
			deptName = "Неизвестное отделение"
		}

		var deptID int
		if err := tx.QueryRow(
			`INSERT INTO departments (name) VALUES ($1)
			 ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
			 RETURNING id`,
			deptName,
		).Scan(&deptID); err != nil {
			return fmt.Errorf("ошибка вставки отделения %s: %v", deptName, err)
		}

		// Группа
		var groupID int
		if err := tx.QueryRow(
			`INSERT INTO groups (department_id, name) VALUES ($1, $2)
			 ON CONFLICT (department_id, name) DO UPDATE SET name = EXCLUDED.name
			 RETURNING id`,
			deptID, g.Group,
		).Scan(&groupID); err != nil {
			return fmt.Errorf("ошибка вставки группы %s: %v", g.Group, err)
		}

		for _, s := range g.Students {
			if s.StudentName == "" {
				continue
			}

			// Студент
			var studentID int
			if err := tx.QueryRow(
				`INSERT INTO students (group_id, full_name) VALUES ($1, $2)
				 ON CONFLICT (group_id, full_name) DO UPDATE SET full_name = EXCLUDED.full_name
				 RETURNING id`,
				groupID, s.StudentName,
			).Scan(&studentID); err != nil {
				return fmt.Errorf("ошибка вставки студента %s: %v", s.StudentName, err)
			}

			// Занятия
			for _, r := range s.Records {
				if r.Date == "" || r.Discipline == "" {
					continue
				}

				// Дата/время — берём как есть, формат у нас строковый (из 1С)
				// Попробуем распарсить в TIMESTAMP
				var dt time.Time
				var parseErr error
				formats := []string{
					"02.01.2006 15:04:05",
					"02.01.2006 0:00:00",
					"02.01.2006",
					time.RFC3339,
				}
				for _, f := range formats {
					dt, parseErr = time.ParseInLocation(f, r.Date, time.Local)
					if parseErr == nil {
						break
					}
				}
				if parseErr != nil {
					log.Printf("[Database] Предупреждение: не удалось распарсить дату занятия %q: %v", r.Date, parseErr)
					continue
				}

				// Вставляем/находим lesson
				var lessonID int
				if err := tx.QueryRow(
					`INSERT INTO lessons (group_id, date_time, discipline, teacher)
					 VALUES ($1, $2, $3, $4)
					 ON CONFLICT (group_id, date_time, discipline)
					 DO UPDATE SET teacher = EXCLUDED.teacher
					 RETURNING id`,
					groupID, dt, r.Discipline, r.Teacher,
				).Scan(&lessonID); err != nil {
					return fmt.Errorf("ошибка вставки занятия (%s, %s): %v", r.Date, r.Discipline, err)
				}

				// Вставляем запись посещаемости по занятию
				if _, err := tx.Exec(
					`INSERT INTO lesson_attendance (lesson_id, student_id, attendance)
					 VALUES ($1, $2, $3)
					 ON CONFLICT (lesson_id, student_id)
					 DO UPDATE SET attendance = EXCLUDED.attendance`,
					lessonID, studentID, r.Attendance,
				); err != nil {
					return fmt.Errorf("ошибка вставки lesson_attendance: %v", err)
				}
			}
		}
	}

	if err := tx.Commit(); err != nil {
		return fmt.Errorf("ошибка коммита транзакции (lessons): %v", err)
	}

	log.Printf("[Database] Расписание занятий загружено. Групп: %d", len(lessonsData.Groups))
	return nil
}
