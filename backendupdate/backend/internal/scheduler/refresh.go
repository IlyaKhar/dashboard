package scheduler

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"dashboard/internal/converter"
)

type Scheduler struct {
	projectRoot        string
	statementInput     string
	statementOutput    string
	studentsInput      string
	studentsOutput     string
	lessonsOutput      string
	scheduleGridInput  string
	scheduleGridOutput string
	pythonScript       string
	// Кэш времени последнего изменения файлов для оптимизации (для сетки расписания)
	lastModified map[string]time.Time
}

func NewScheduler(projectRoot, attendanceInput, attendanceOutput, statementInput, statementOutput, studentsInput, studentsOutput, lessonsInput, lessonsOutput, scheduleGridInput, scheduleGridOutput, pythonScript string) *Scheduler {
	return &Scheduler{
		projectRoot:        projectRoot,
		statementInput:     statementInput,
		statementOutput:    statementOutput,
		studentsInput:      studentsInput,
		studentsOutput:     studentsOutput,
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

	// Единственный мастер‑файл: ведомость (ведомость.xls / ведомость.xlsx)
	if _, err := os.Stat(s.statementInput); err != nil {
		log.Printf("[Scheduler] Файл ведомости не найден: %s", s.statementInput)
		return err
	}

	log.Printf("[Scheduler] Найден файл ведомости: %s", s.statementInput)
	log.Println("[Scheduler] Запуск мастер-конвертера (ведомость → все JSON)...")

	outputDir := filepath.Dir(s.statementOutput) // обычно public/
	result, err := converter.ConvertMaster(s.statementInput, outputDir, s.pythonScript)
	if err != nil {
		log.Printf("[Scheduler] Ошибка мастер-конвертации: %v", err)
		return err
	}

	if result.StudentsOutput != "" {
		log.Printf("[Scheduler] ✓ students.json создан: %s", result.StudentsOutput)
	} else {
		log.Println("[Scheduler] Предупреждение: мастер-конвертер не вернул путь к students.json")
	}
	if result.AttendanceOutput != "" {
		log.Printf("[Scheduler] ✓ attendance.json создан: %s", result.AttendanceOutput)
	} else {
		log.Println("[Scheduler] Предупреждение: мастер-конвертер не вернул путь к attendance.json")
	}
	if result.VedomostOutput != "" {
		log.Printf("[Scheduler] ✓ vedomost.json создан: %s", result.VedomostOutput)
	}

	if len(result.Warnings) > 0 {
		for _, w := range result.Warnings {
			log.Printf("[Scheduler] Предупреждение мастер-конвертера: %s", w)
		}
	}
	if len(result.Errors) > 0 {
		for _, e := range result.Errors {
			log.Printf("[Scheduler] Ошибка мастер-конвертера: %s", e)
		}
	}

	// Гарантируем наличие vedomost.json, чтобы сервисы и jsonstore не падали с "файл не найден".
	// Если мастер-конвертер не смог собрать сводную ведомость, создаём пустой файл-«заглушку».
	if _, err := os.Stat(s.statementOutput); os.IsNotExist(err) {
		log.Printf("[Scheduler] vedomost.json не найден, создаём пустой файл: %s", s.statementOutput)
		if writeErr := os.WriteFile(s.statementOutput, []byte("[]"), 0o644); writeErr != nil {
			log.Printf("[Scheduler] Не удалось создать пустой vedomost.json: %v", writeErr)
		}
	}

	// Если есть отдельный файл контингента (студенты.xlsx), поверх данных из ведомости
	// пересобираем students.json через специализированный конвертер.
	if s.studentsInput != "" {
		if _, err := os.Stat(s.studentsInput); err == nil {
			log.Printf("[Scheduler] Обнаружен отдельный файл контингента: %s", s.studentsInput)
			log.Println("[Scheduler] Пересборка students.json из файла контингента...")
			if err := converter.ConvertStudents(s.studentsInput, s.studentsOutput); err != nil {
				log.Printf("[Scheduler] Ошибка конвертации контингента студентов: %v", err)
			} else {
				log.Printf("[Scheduler] ✓ students.json обновлён из файла контингента: %s", s.studentsOutput)
			}
		} else if !os.IsNotExist(err) {
			log.Printf("[Scheduler] Не удалось проверить файл контингента %s: %v", s.studentsInput, err)
		}
	}

	// Проверяем наличие файла сетки расписания (расписание.xls) и его изменения
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
					log.Println("[Scheduler] Используем ранее сформированное расписание (schedule.json)")
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
// Оставлен для совместимости с тестами (TestScheduler_shouldUpdateFile).
// Возвращает true, если входной файл новее выходного или выходного файла нет.
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
