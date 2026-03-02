package scheduler

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"dashboard/internal/converter"
)

type Scheduler struct {
	projectRoot         string
	attendanceInput     string
	attendanceOutput    string
	statementInput      string
	statementOutput     string
	studentsInput       string
	studentsOutput      string
	lessonsInput        string
	lessonsOutput       string
	scheduleGridInput   string
	scheduleGridOutput  string
	pythonScript        string
	// Кэш времени последнего изменения файлов для оптимизации
	lastModified map[string]time.Time
}

func NewScheduler(projectRoot, attendanceInput, attendanceOutput, statementInput, statementOutput, studentsInput, studentsOutput, lessonsInput, lessonsOutput, scheduleGridInput, scheduleGridOutput, pythonScript string) *Scheduler {
	return &Scheduler{
		projectRoot:        projectRoot,
		attendanceInput:    attendanceInput,
		attendanceOutput:   attendanceOutput,
		statementInput:     statementInput,
		statementOutput:    statementOutput,
		studentsInput:      studentsInput,
		studentsOutput:     studentsOutput,
		lessonsInput:       lessonsInput,
		lessonsOutput:      lessonsOutput,
		scheduleGridInput:  scheduleGridInput,
		scheduleGridOutput: scheduleGridOutput,
		pythonScript:       pythonScript,
		lastModified:       make(map[string]time.Time),
	}
}

// RefreshData обновляет данные, запуская оба конвертера
// Проверяет изменения файлов перед конвертацией (оптимизация)
// Приоритет: если есть ведомость.xls, используем мастер-конвертер для создания всех JSON
func (s *Scheduler) RefreshData() error {
	log.Println("[Scheduler] Начало обновления данных...")

	// Проверяем, есть ли ведомость.xls - если да, используем мастер-конвертер
	statementFile := s.statementInput
	if _, err := os.Stat(statementFile); err != nil {
		// Пробуем найти ведомость.xlsx
		statementFileXLSX := strings.TrimSuffix(statementFile, ".xls") + ".xlsx"
		if _, err2 := os.Stat(statementFileXLSX); err2 == nil {
			statementFile = statementFileXLSX
		}
	}

	// Переменные для отслеживания, что создал мастер-конвертер
	needAttendance := true
	needStudents := true
	needVedomost := true

	// Если ведомость найдена, используем мастер-конвертер
	if _, err := os.Stat(statementFile); err == nil {
		log.Printf("[Scheduler] Найден файл ведомости: %s", statementFile)
		log.Println("[Scheduler] Используем мастер-конвертер для создания всех JSON...")
		
		outputDir := filepath.Dir(s.statementOutput) // public/
		result, err := converter.ConvertMaster(statementFile, outputDir, s.pythonScript)
		if err != nil {
			log.Printf("[Scheduler] Ошибка мастер-конвертации: %v", err)
			log.Println("[Scheduler] Пробуем использовать отдельные конвертеры...")
		} else {
			// Мастер-конвертер успешно создал файлы
			if result.StudentsOutput != "" {
				log.Printf("[Scheduler] ✓ students.json создан: %s", result.StudentsOutput)
				needStudents = false
			}
			if result.AttendanceOutput != "" {
				log.Printf("[Scheduler] ✓ attendance.json создан: %s", result.AttendanceOutput)
				needAttendance = false
			}
			if result.VedomostOutput != "" {
				log.Printf("[Scheduler] ✓ vedomost.json создан: %s", result.VedomostOutput)
				needVedomost = false
			}
			if len(result.Warnings) > 0 {
				for _, w := range result.Warnings {
					log.Printf("[Scheduler] Предупреждение: %s", w)
				}
			}
			if len(result.Errors) > 0 {
				for _, e := range result.Errors {
					log.Printf("[Scheduler] Ошибка: %s", e)
				}
			}
			// Обновляем время последнего изменения
			if info, err := os.Stat(statementFile); err == nil {
				s.lastModified[statementFile] = info.ModTime()
			}
			log.Println("[Scheduler] Мастер-конвертация завершена, переходим к проверке остальных файлов...")
		}
	}

	// Проверяем наличие входных файлов и их изменения (fallback на отдельные файлы)
	// Если мастер-конвертер не создал какой-то файл, запускаем отдельный конвертер
	
	// Проверяем наличие файла посещаемости
	if needAttendance {
		// Мастер-конвертер не создал файл, всегда запускаем отдельный конвертер
		if _, err := os.Stat(s.attendanceInput); err == nil {
			log.Println("[Scheduler] Конвертация посещаемости (мастер-конвертер не создал файл)...")
			if err := converter.ConvertAttendance(s.attendanceInput, s.attendanceOutput); err != nil {
				log.Printf("[Scheduler] Ошибка конвертации посещаемости: %v", err)
				// Не возвращаем ошибку, продолжаем работу
			} else {
				// Обновляем время последнего изменения
				if info, err := os.Stat(s.attendanceInput); err == nil {
					s.lastModified[s.attendanceInput] = info.ModTime()
				}
				log.Println("[Scheduler] Посещаемость обновлена")
			}
		} else {
			log.Printf("[Scheduler] Входной файл посещаемости не найден: %s", s.attendanceInput)
		}
	} else {
		log.Println("[Scheduler] Посещаемость уже создана мастер-конвертером, пропускаем")
	}

	// Проверяем наличие файла ведомости и его изменения
	if needVedomost {
		if shouldUpdate, err := s.shouldUpdateFile(s.statementInput, s.statementOutput); err != nil {
			log.Printf("[Scheduler] Предупреждение: %v", err)
		} else if shouldUpdate || !fileExists(s.statementOutput) {
			// Конвертируем ведомость
			log.Println("[Scheduler] Конвертация ведомости...")
			if err := converter.ConvertStatement(s.statementInput, s.statementOutput, s.pythonScript); err != nil {
				log.Printf("[Scheduler] Ошибка конвертации ведомости: %v", err)
				// Не возвращаем ошибку, продолжаем работу
			} else {
				// Обновляем время последнего изменения
				if info, err := os.Stat(s.statementInput); err == nil {
					s.lastModified[s.statementInput] = info.ModTime()
				}
				log.Println("[Scheduler] Ведомость обновлена")
			}
		} else {
			log.Println("[Scheduler] Ведомость не изменилась, пропускаем")
		}
	} else {
		log.Println("[Scheduler] Ведомость уже создана мастер-конвертером, пропускаем")
	}

	// Проверяем наличие файла контингента студентов и его изменения
	if needStudents {
		log.Printf("[Scheduler] Проверка файла студентов: входной=%s, выходной=%s", s.studentsInput, s.studentsOutput)
		// Мастер-конвертер не создал файл, всегда запускаем отдельный конвертер
		if _, err := os.Stat(s.studentsInput); err == nil {
			log.Println("[Scheduler] Конвертация контингента студентов (мастер-конвертер не создал файл)...")
			if err := converter.ConvertStudents(s.studentsInput, s.studentsOutput); err != nil {
				log.Printf("[Scheduler] Ошибка конвертации контингента студентов: %v", err)
				// Не возвращаем ошибку, чтобы не ломать остальные конвертеры
			} else {
				// Обновляем время последнего изменения
				if info, err := os.Stat(s.studentsInput); err == nil {
					s.lastModified[s.studentsInput] = info.ModTime()
				}
				log.Println("[Scheduler] Контингент студентов обновлён")
			}
		} else {
			log.Printf("[Scheduler] Входной файл контингента студентов не найден: %s", s.studentsInput)
		}
	} else {
		log.Println("[Scheduler] Контингент студентов уже создан мастер-конвертером, пропускаем")
	}

	// Проверяем наличие файла расписания занятий и его изменения
	// Всегда обновляем расписание занятий при каждом запуске
	if _, err := os.Stat(s.lessonsInput); err == nil {
		log.Println("[Scheduler] Конвертация расписания занятий...")
		if err := converter.ConvertLessons(s.lessonsInput, s.lessonsOutput); err != nil {
			log.Printf("[Scheduler] Ошибка конвертации расписания занятий: %v", err)
			// Не возвращаем ошибку, чтобы не ломать остальные конвертеры
		} else {
			// Обновляем время последнего изменения
			if info, err := os.Stat(s.lessonsInput); err == nil {
				s.lastModified[s.lessonsInput] = info.ModTime()
			}
			log.Println("[Scheduler] Расписание занятий обновлено")
		}
	} else {
		log.Printf("[Scheduler] Входной файл расписания занятий не найден: %s", s.lessonsInput)
	}

	// Проверяем наличие файла сетки расписания (расписание.xls) и его изменения
	// Всегда обновляем сетку расписания при каждом запуске
	if _, err := os.Stat(s.scheduleGridInput); err == nil {
		log.Println("[Scheduler] Конвертация сетки расписания (расписание.xls)...")
		
		// Сначала конвертируем в формат сетки
		gridOutput := s.scheduleGridOutput
		if err := converter.ConvertScheduleGrid(s.scheduleGridInput, gridOutput, ""); err != nil {
			log.Printf("[Scheduler] Ошибка конвертации сетки расписания: %v", err)
		} else {
			log.Println("[Scheduler] Сетка расписания создана")
			
			// Затем преобразуем в формат lessons.json для совместимости
			// Используем сетку как основной источник, если она есть
			if _, err := os.Stat(s.studentsOutput); err == nil {
				log.Println("[Scheduler] Преобразование сетки расписания в формат lessons...")
				if err := converter.ConvertScheduleGridToLessonsFormat(
					gridOutput,
					s.studentsOutput,
					s.lessonsOutput,
					"",
				); err != nil {
					log.Printf("[Scheduler] Ошибка преобразования сетки в формат lessons: %v", err)
					log.Println("[Scheduler] Используем старое расписание из Проба.xlsx")
				} else {
					log.Println("[Scheduler] Расписание обновлено из сетки расписания")
					// Обновляем время последнего изменения
					if info, err := os.Stat(s.scheduleGridInput); err == nil {
						s.lastModified[s.scheduleGridInput] = info.ModTime()
					}
				}
			}
		}
	} else {
		log.Printf("[Scheduler] Входной файл сетки расписания не найден: %s", s.scheduleGridInput)
	}

	log.Println("[Scheduler] Обновление данных завершено успешно!")
	return nil
}

// shouldUpdateFile проверяет, нужно ли обновлять файл
// Возвращает true, если входной файл новее выходного или выходного файла нет
func (s *Scheduler) shouldUpdateFile(inputFile, outputFile string) (bool, error) {
	// Проверяем наличие входного файла
	inputInfo, err := os.Stat(inputFile)
	if os.IsNotExist(err) {
		return false, fmt.Errorf("входной файл не найден: %s", inputFile)
	}
	if err != nil {
		return false, fmt.Errorf("ошибка проверки входного файла: %v", err)
	}

	// Проверяем наличие выходного файла
	outputInfo, err := os.Stat(outputFile)
	if os.IsNotExist(err) {
		// Выходного файла нет - нужно обновить
		return true, nil
	}
	if err != nil {
		return false, fmt.Errorf("ошибка проверки выходного файла: %v", err)
	}

	// Сравниваем время изменения
	// Если входной файл новее выходного - нужно обновить
	if inputInfo.ModTime().After(outputInfo.ModTime()) {
		return true, nil
	}

	// Проверяем кэш (если файл уже обрабатывался в этой сессии)
	if lastMod, exists := s.lastModified[inputFile]; exists {
		if inputInfo.ModTime().After(lastMod) {
			return true, nil
		}
	}

	return false, nil
}

// fileExists проверяет, существует ли файл
func fileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}
