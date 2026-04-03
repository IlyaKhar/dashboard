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
	projectRoot     string
	statementInput  string
	statementOutput string
	studentsInput   string
	studentsOutput  string
	lessonsOutput   string
	// РасписаниеНаДату.xls — ежедневное расписание с шары
	scheduleGridInput string
	pythonScript      string
	lastModified      map[string]time.Time
}

func NewScheduler(projectRoot, attendanceInput, attendanceOutput, statementInput, statementOutput, studentsInput, studentsOutput, lessonsInput, lessonsOutput, scheduleGridInput, scheduleGridOutput, pythonScript string) *Scheduler {
	return &Scheduler{
		projectRoot:       projectRoot,
		statementInput:    statementInput,
		statementOutput:   statementOutput,
		studentsInput:     studentsInput,
		studentsOutput:    studentsOutput,
		lessonsOutput:     lessonsOutput,
		scheduleGridInput: scheduleGridInput,
		pythonScript:      pythonScript,
		lastModified:      make(map[string]time.Time),
	}
}

// RefreshData обновляет данные:
// 1. Ведомость → attendance.json, students.json, vedomost.json, attendance_history.json
// 2. РасписаниеНаДату.xls → schedule.json (сегодня) + schedule_history.json (накопление)
func (s *Scheduler) RefreshData() error {
	log.Println("[Scheduler] Начало обновления данных...")

	// Единственный мастер‑файл: ведомость (ведомость.xls / ведомость.xlsx)
	if _, err := os.Stat(s.statementInput); err != nil {
		log.Printf("[Scheduler] Файл ведомости не найден: %s", s.statementInput)
		return err
	}

	log.Printf("[Scheduler] Найден файл ведомости: %s", s.statementInput)
	log.Println("[Scheduler] Запуск мастер-конвертера (ведомость → все JSON)...")

	outputDir := filepath.Dir(s.statementOutput)
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

	// Гарантируем наличие vedomost.json
	if _, err := os.Stat(s.statementOutput); os.IsNotExist(err) {
		log.Printf("[Scheduler] vedomost.json не найден, создаём пустой файл: %s", s.statementOutput)
		if writeErr := os.WriteFile(s.statementOutput, []byte("[]"), 0o644); writeErr != nil {
			log.Printf("[Scheduler] Не удалось создать пустой vedomost.json: %v", writeErr)
		}
	}

	// Если есть отдельный файл контингента (студенты.xlsx)
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

	// Обработка ежедневного расписания (РасписаниеНаДату.xls)
	// → schedule.json (только сегодня, оперативный)
	// → schedule_history.json (накопление, исторический)
	if _, err := os.Stat(s.scheduleGridInput); err == nil {
		log.Printf("[Scheduler] Обнаружен файл ежедневного расписания: %s", s.scheduleGridInput)
		if _, err := os.Stat(s.studentsOutput); err == nil {
			if err := converter.ConvertDailySchedule(
				s.scheduleGridInput,
				s.lessonsOutput,
				s.studentsOutput,
				s.pythonScript,
			); err != nil {
				log.Printf("[Scheduler] Ошибка конвертации ежедневного расписания: %v", err)
			} else {
				log.Println("[Scheduler] ✓ schedule.json (сегодня) + schedule_history.json (накопление) обновлены")
				if info, err := os.Stat(s.scheduleGridInput); err == nil {
					s.lastModified[s.scheduleGridInput] = info.ModTime()
				}
			}
		} else {
			log.Println("[Scheduler] students.json не найден, пропускаем конвертацию расписания")
		}
	} else {
		log.Printf("[Scheduler] Входной файл расписания не найден: %s", s.scheduleGridInput)
	}

	log.Println("[Scheduler] Обновление данных завершено успешно!")
	return nil
}

// shouldUpdateFile проверяет, нужно ли обновлять файл
func (s *Scheduler) shouldUpdateFile(inputFile, outputFile string) (bool, error) {
	inputInfo, err := os.Stat(inputFile)
	if os.IsNotExist(err) {
		return false, fmt.Errorf("входной файл не найден: %s", inputFile)
	}
	if err != nil {
		return false, fmt.Errorf("ошибка проверки входного файла: %v", err)
	}

	outputInfo, err := os.Stat(outputFile)
	if os.IsNotExist(err) {
		return true, nil
	}
	if err != nil {
		return false, fmt.Errorf("ошибка проверки выходного файла: %v", err)
	}

	if inputInfo.ModTime().After(outputInfo.ModTime()) {
		return true, nil
	}

	if lastMod, exists := s.lastModified[inputFile]; exists {
		if inputInfo.ModTime().After(lastMod) {
			return true, nil
		}
	}

	return false, nil
}
