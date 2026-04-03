package config

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

// Config содержит конфигурацию приложения
type Config struct {
	// Интервал обновления данных
	RefreshInterval time.Duration

	// Пути к файлам
	ProjectRoot        string
	AttendanceInput    string
	AttendanceOutput   string
	StatementInput     string
	StatementOutput    string // public/vedomost.json
	StudentsInput      string // Ведомостьколва.xlsx или Контингент студентов
	StudentsOutput     string // public/students.json
	LessonsInput       string // (не используется, расписание теперь строится из сетки/ОКЭИ)
	LessonsOutput      string // public/schedule.json
	ScheduleGridInput  string // РасписаниеНаДату.xls — ежедневное расписание с шары
	ScheduleGridOutput string // (не используется, оставлен для совместимости)
	PythonScript       string

	// Расписание с сайта ОКЭИ (timetable: группа → день → пара → дисциплина)
	OkseiScheduleURL    string // URL страницы расписания
	OkseiScheduleOutput string // путь к JSON с расписанием ОКЭИ

	// Интеграция с 1С
	OneCSourceDir string // Путь к директории, куда смонтирован \\1C01\proba

	// Настройки сервера
	ServerPort string
	ServerHost string

	// Настройки БД
	DatabaseURL      string
	DatabaseHost     string
	DatabasePort     string
	DatabaseUser     string
	DatabasePassword string
	DatabaseName     string

	// JWT авторизация (из attendance-backend)
	JWTSecret string

	// CORS (из attendance-backend)
	CORSOrigins []string

	// Алерты (из attendance-backend)
	AbsenceThreshold int

	// Логин (из attendance-backend)
	LoginUser     string
	LoginPassword string
	LoginRole     string
}

// Load загружает конфигурацию из переменных окружения или использует значения по умолчанию
func Load() (*Config, error) {
	// Получаем корневую директорию проекта
	// Используем рабочую директорию при запуске сервера
	wd, err := os.Getwd()
	if err != nil {
		return nil, fmt.Errorf("ошибка получения рабочей директории: %v", err)
	}

	var projectRoot string
	// Если запускаем из backend/, поднимаемся на уровень выше
	if filepath.Base(wd) == "backend" {
		projectRoot = filepath.Dir(wd)
	} else {
		// Ищем директорию с папкой public/ и backend/
		current := wd
		for {
			// Проверяем наличие папок public/ и backend/ в текущей директории
			publicExists := false
			backendExists := false

			if _, err := os.Stat(filepath.Join(current, "public")); err == nil {
				publicExists = true
			}
			if _, err := os.Stat(filepath.Join(current, "backend")); err == nil {
				backendExists = true
			}

			// Если есть обе папки - это корень проекта
			if publicExists && backendExists {
				projectRoot = current
				break
			}

			// Поднимаемся на уровень выше
			parent := filepath.Dir(current)
			if parent == current || parent == "/" {
				// Дошли до корня, не нашли
				break
			}
			current = parent
		}

		// Если не нашли, пробуем найти по наличию public/
		if projectRoot == "" {
			current := wd
			for {
				if _, err := os.Stat(filepath.Join(current, "public")); err == nil {
					projectRoot = current
					break
				}
				parent := filepath.Dir(current)
				if parent == current || parent == "/" {
					break
				}
				current = parent
			}
		}

		// Если всё ещё не нашли, используем текущую директорию
		if projectRoot == "" {
			projectRoot = wd
		}
	}

	// Проверяем, что директория существует и содержит public/
	if _, err := os.Stat(filepath.Join(projectRoot, "public")); os.IsNotExist(err) {
		return nil, fmt.Errorf("не найдена директория проекта (ожидается папка 'public' в %s)", projectRoot)
	}

	// Интервал обновления (по умолчанию 90 минут)
	refreshInterval := 90 * time.Minute
	if intervalStr := os.Getenv("REFRESH_INTERVAL"); intervalStr != "" {
		if parsed, err := time.ParseDuration(intervalStr); err == nil {
			refreshInterval = parsed
		}
	}

	// Порт сервера (по умолчанию 8080)
	serverPort := os.Getenv("SERVER_PORT")
	if serverPort == "" {
		serverPort = "8080"
	}

	// Хост сервера (по умолчанию localhost)
	serverHost := os.Getenv("SERVER_HOST")
	if serverHost == "" {
		serverHost = "localhost"
	}

	// Настройки БД
	databaseURL := os.Getenv("DATABASE_URL")
	if databaseURL == "" {
		// Формируем URL из отдельных параметров, если DATABASE_URL не указан
		dbHost := os.Getenv("DB_HOST")
		if dbHost == "" {
			dbHost = "localhost"
		}
		dbPort := os.Getenv("DB_PORT")
		if dbPort == "" {
			dbPort = "5432"
		}
		dbUser := os.Getenv("DB_USER")
		if dbUser == "" {
			dbUser = "postgres"
		}
		dbPassword := os.Getenv("DB_PASSWORD")
		dbName := os.Getenv("DB_NAME")
		if dbName == "" {
			dbName = "dashboard"
		}

		if dbPassword != "" {
			databaseURL = fmt.Sprintf("postgres://%s:%s@%s:%s/%s?sslmode=disable",
				dbUser, dbPassword, dbHost, dbPort, dbName)
		}
	}

	// JWT Secret — в production обязан быть задан и отличаться от дефолта
	jwtSecret := os.Getenv("JWT_SECRET")
	if jwtSecret == "" {
		jwtSecret = "change-me-in-production"
	}
	if os.Getenv("APP_ENV") == "production" && jwtSecret == "change-me-in-production" {
		return nil, fmt.Errorf("в production JWT_SECRET обязан быть задан и отличаться от дефолта")
	}

	// CORS Origins (из attendance-backend)
	corsEnv := strings.TrimSpace(os.Getenv("CORS_ORIGINS"))
	var corsOrigins []string
	if corsEnv == "" || corsEnv == "*" {
		corsOrigins = []string{"*"}
	} else {
		parts := strings.Split(corsEnv, ",")
		trimmed := make([]string, 0, len(parts))
		for _, o := range parts {
			s := strings.TrimSpace(o)
			if s != "" {
				trimmed = append(trimmed, s)
			}
		}
		if len(trimmed) == 0 {
			trimmed = []string{"http://localhost:3000", "http://localhost:5173"}
		}
		corsOrigins = trimmed
	}

	// Absence Threshold (из attendance-backend)
	threshold, _ := strconv.Atoi(os.Getenv("ABSENCE_THRESHOLD"))
	if threshold <= 0 || threshold > 100 {
		threshold = 10
	}

	// Login credentials — в production пароль обязан быть задан и отличаться от admin.
	// LOGIN_PASSWORD в production ожидается как bcrypt-хэш.
	loginUser := os.Getenv("LOGIN_USER")
	if loginUser == "" {
		loginUser = "admin"
	}
	loginPassword := os.Getenv("LOGIN_PASSWORD")
	if loginPassword == "" {
		loginPassword = "admin"
	}
	if os.Getenv("APP_ENV") == "production" && loginPassword == "admin" {
		return nil, fmt.Errorf("в production LOGIN_PASSWORD обязан быть задан и отличаться от admin")
	}
	loginRole := os.Getenv("LOGIN_ROLE")
	if loginRole == "" {
		loginRole = "admin"
	}

	// Путь к исходным файлам 1С (шаре \\1C01\proba)
	oneCSourceDir := os.Getenv("ONEC_SOURCE_DIR")
	if oneCSourceDir == "" {
		oneCSourceDir = `\\1C01\proba`
	}

	// Пути к файлам — можно переопределить через env
	// Ищем файлы в нескольких местах: корень проекта, converter/, backend/internal/converter/
	converterDir := filepath.Join(projectRoot, "backend", "internal", "converter")

	// Функция для поиска файла в нескольких местах
	findFile := func(filename string) string {
		// Проверяем переменную окружения
		if envPath := os.Getenv("ATTENDANCE_INPUT"); filename == "Посещаемость.xlsx" && envPath != "" {
			return envPath
		}
		if envPath := os.Getenv("STATEMENT_INPUT"); filename == "ведомость.xls" && envPath != "" {
			return envPath
		}
		if envPath := os.Getenv("STUDENTS_INPUT"); (filename == "Ведомостьколва.xlsx" || filename == "ведомостьколва.xlsx") && envPath != "" {
			return envPath
		}
		if envPath := os.Getenv("LESSONS_INPUT"); filename == "Проба.xlsx" && envPath != "" {
			return envPath
		}
		if envPath := os.Getenv("SCHEDULE_GRID_INPUT"); filename == "РасписаниеНаДату.xls" && envPath != "" {
			return envPath
		}

		// Ищем в разных местах
		locations := []string{
			filepath.Join(projectRoot, filename),
			filepath.Join(converterDir, filename),
			filepath.Join(projectRoot, "converter", filename),
		}

		for _, loc := range locations {
			if _, err := os.Stat(loc); err == nil {
				return loc
			}
		}

		// По умолчанию ожидаем входные Excel в директории конвертера.
		// Это важно для сценария, когда файлы синхронизируются с 1С уже после старта сервера.
		return filepath.Join(converterDir, filename)
	}

	attendanceInput := os.Getenv("ATTENDANCE_INPUT")
	if attendanceInput == "" {
		attendanceInput = findFile("Посещаемость.xlsx")
	}
	attendanceOutput := os.Getenv("ATTENDANCE_OUTPUT")
	if attendanceOutput == "" {
		attendanceOutput = filepath.Join(projectRoot, "public", "attendance.json")
	}

	statementInput := os.Getenv("STATEMENT_INPUT")
	if statementInput == "" {
		// Ищем оба варианта ведомости и выбираем наиболее новый файл (по времени изменения).
		xlsPath := findFile("ведомость.xls")
		xlsxPath := findFile("ведомость.xlsx")

		var xlsInfo, xlsxInfo os.FileInfo
		var xlsErr, xlsxErr error

		if xlsPath != "" {
			xlsInfo, xlsErr = os.Stat(xlsPath)
		}
		if xlsxPath != "" {
			xlsxInfo, xlsxErr = os.Stat(xlsxPath)
		}

		switch {
		case xlsErr == nil && xlsxErr == nil:
			// Оба файла есть — берём тот, что новее.
			if xlsxInfo.ModTime().After(xlsInfo.ModTime()) {
				statementInput = xlsxPath
			} else {
				statementInput = xlsPath
			}
		case xlsErr == nil:
			// Есть только .xls
			statementInput = xlsPath
		case xlsxErr == nil:
			// Есть только .xlsx
			statementInput = xlsxPath
		default:
			// Ничего не нашли — по умолчанию ожидаем ведомость в директории конвертера.gg
			statementInput = filepath.Join(converterDir, "ведомость.xls")
		} 
	}
	statementOutput := os.Getenv("STATEMENT_OUTPUT")
	if statementOutput == "" {
		// Ведомость по пропускам
		statementOutput = filepath.Join(projectRoot, "public", "vedomost.json")
	}

	studentsInput := os.Getenv("STUDENTS_INPUT")
	if studentsInput == "" {
		// Приоритетно используем новый файл контингента "студенты.xlsx" (копируется с шары).
		// Если его нет — откатываемся к старым вариантам (Ведомостьколва.xlsx и т.п.).
		studentsInput = findFile("студенты.xlsx")
		if _, err := os.Stat(studentsInput); err != nil {
			studentsInput = findFile("Ведомостьколва.xlsx")
			if _, err := os.Stat(studentsInput); err != nil {
				studentsInput = findFile("ведомостьколва.xlsx")
			}
		}
	}
	studentsOutput := os.Getenv("STUDENTS_OUTPUT")
	if studentsOutput == "" {
		studentsOutput = filepath.Join(projectRoot, "public", "students.json")
	}

	// Входной Excel для расписания напрямую больше не используем —
	// schedule.json формируется из сетки расписания (расписание.xls) и данных студентов.
	lessonsInput := os.Getenv("LESSONS_INPUT")
	lessonsOutput := os.Getenv("LESSONS_OUTPUT")
	if lessonsOutput == "" {
		// Нормализованное расписание
		lessonsOutput = filepath.Join(projectRoot, "public", "schedule.json")
	}

	scheduleGridInput := os.Getenv("SCHEDULE_GRID_INPUT")
	if scheduleGridInput == "" {
		// РасписаниеНаДату.xls — ежедневное расписание с шары (обновляется каждый день)
		scheduleGridInput = findFile("РасписаниеНаДату.xls")
	}
	scheduleGridOutput := os.Getenv("SCHEDULE_GRID_OUTPUT")
	if scheduleGridOutput == "" {
		scheduleGridOutput = filepath.Join(projectRoot, "public", "schedule_grid.json")
	}

	// Расписание с сайта ОКЭИ
	okseiScheduleURL := getOkseiScheduleURL()
	okseiScheduleOutput := os.Getenv("OKSEI_SCHEDULE_OUTPUT")
	if okseiScheduleOutput == "" {
		okseiScheduleOutput = filepath.Join(projectRoot, "public", "oksei_timetable.json")
	}

	pythonScript := os.Getenv("PYTHON_SCRIPT")
	if pythonScript == "" {
		// Ищем Python скрипт в нескольких местах
		possiblePaths := []string{
			filepath.Join(converterDir, "xls_to_xlsx.py"),
			filepath.Join(projectRoot, "backend", "internal", "converter", "xls_to_xlsx.py"),
			filepath.Join(projectRoot, "statement-converter", "xls_to_xlsx.py"),
		}
		pythonScript = possiblePaths[0] // По умолчанию первый путь
		for _, path := range possiblePaths {
			if _, err := os.Stat(path); err == nil {
				pythonScript = path
				break
			}
		}
	}

	cfg := &Config{
		RefreshInterval:     refreshInterval,
		ProjectRoot:         projectRoot,
		AttendanceInput:     attendanceInput,
		AttendanceOutput:    attendanceOutput,
		StatementInput:      statementInput,
		StatementOutput:     statementOutput,
		StudentsInput:       studentsInput,
		StudentsOutput:      studentsOutput,
		LessonsInput:        lessonsInput,
		LessonsOutput:       lessonsOutput,
		ScheduleGridInput:   scheduleGridInput,
		ScheduleGridOutput:  scheduleGridOutput,
		PythonScript:        pythonScript,
		OkseiScheduleURL:    okseiScheduleURL,
		OkseiScheduleOutput: okseiScheduleOutput,
		OneCSourceDir:       oneCSourceDir,
		ServerPort:          serverPort,
		ServerHost:          serverHost,
		DatabaseURL:         databaseURL,
		DatabaseHost:        os.Getenv("DB_HOST"),
		DatabasePort:        os.Getenv("DB_PORT"),
		DatabaseUser:        os.Getenv("DB_USER"),
		DatabasePassword:    os.Getenv("DB_PASSWORD"),
		DatabaseName:        os.Getenv("DB_NAME"),
		JWTSecret:           jwtSecret,
		CORSOrigins:         corsOrigins,
		AbsenceThreshold:    threshold,
		LoginUser:           loginUser,
		LoginPassword:       loginPassword,
		LoginRole:           loginRole,
	}

	return cfg, nil
}

// getOkseiScheduleURL возвращает URL страницы расписания ОКЭИ (или прямой ссылки на файл).
func getOkseiScheduleURL() string {
	if u := strings.TrimSpace(os.Getenv("OKSEI_SCHEDULE_URL")); u != "" {
		return u
	}
	return "https://oksei.ru/studentu/raspisanie_uchebnykh_zanyatij"
}
