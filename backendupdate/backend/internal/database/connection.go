package database

import (
	"database/sql"
	"fmt"
	"log"
	"os"
	"path/filepath"

	_ "github.com/lib/pq" // PostgreSQL драйвер
)

// DB содержит подключение к базе данных
var DB *sql.DB

// Connect подключается к PostgreSQL
func Connect(databaseURL string) error {
	if databaseURL == "" {
		return fmt.Errorf("DATABASE_URL не указан")
	}

	var err error
	DB, err = sql.Open("postgres", databaseURL)
	if err != nil {
		return fmt.Errorf("ошибка подключения к БД: %v", err)
	}

	// Проверяем подключение
	if err := DB.Ping(); err != nil {
		return fmt.Errorf("ошибка ping БД: %v", err)
	}

	log.Println("[Database] Подключение к PostgreSQL установлено")
	return nil
}

// Close закрывает подключение к БД
func Close() error {
	if DB != nil {
		return DB.Close()
	}
	return nil
}

// InitSchema создаёт схему БД из SQL файла
func InitSchema() error {
	if DB == nil {
		return fmt.Errorf("БД не подключена")
	}

	// Пробуем разные пути к schema.sql
	possiblePaths := []string{
		filepath.Join("internal", "database", "schema.sql"),
		filepath.Join("backend", "internal", "database", "schema.sql"),
		"schema.sql",
		filepath.Join("..", "internal", "database", "schema.sql"),
	}

	var sqlBytes []byte
	var err error
	for _, schemaPath := range possiblePaths {
		sqlBytes, err = os.ReadFile(schemaPath)
		if err == nil {
			log.Printf("[Database] Схема найдена: %s", schemaPath)
			break
		}
	}

	if err != nil {
		// Если файл не найден, используем встроенную схему
		log.Println("[Database] Файл schema.sql не найден, используем встроенную схему")
		sqlBytes = []byte(getEmbeddedSchema())
	}

	// Выполняем SQL
	if _, err := DB.Exec(string(sqlBytes)); err != nil {
		return fmt.Errorf("ошибка выполнения schema.sql: %v", err)
	}

	log.Println("[Database] Схема БД инициализирована")
	return nil
}

// getEmbeddedSchema возвращает встроенную SQL схему (на случай, если файл не найден)
func getEmbeddedSchema() string {
	return `
-- Встроенная схема БД
CREATE TABLE IF NOT EXISTS departments (
    id SERIAL PRIMARY KEY,
    name VARCHAR(255) NOT NULL UNIQUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS groups (
    id SERIAL PRIMARY KEY,
    department_id INTEGER NOT NULL REFERENCES departments(id) ON DELETE CASCADE,
    name VARCHAR(50) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(department_id, name)
);

CREATE TABLE IF NOT EXISTS students (
    id SERIAL PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES groups(id) ON DELETE CASCADE,
    full_name VARCHAR(255) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(group_id, full_name)
);

CREATE TABLE IF NOT EXISTS attendance (
    id SERIAL PRIMARY KEY,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    date DATE NOT NULL,
    missed_hours INTEGER NOT NULL DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(student_id, date)
);

CREATE TABLE IF NOT EXISTS specialties (
    id SERIAL PRIMARY KEY,
    department_id INTEGER NOT NULL REFERENCES departments(id) ON DELETE CASCADE,
    name VARCHAR(255) NOT NULL,
    total_missed INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(department_id, name)
);

CREATE TABLE IF NOT EXISTS summary_groups (
    id SERIAL PRIMARY KEY,
    specialty_id INTEGER NOT NULL REFERENCES specialties(id) ON DELETE CASCADE,
    name VARCHAR(50) NOT NULL,
    total_missed INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(specialty_id, name)
);

CREATE TABLE IF NOT EXISTS summary_students (
    id SERIAL PRIMARY KEY,
    summary_group_id INTEGER NOT NULL REFERENCES summary_groups(id) ON DELETE CASCADE,
    full_name VARCHAR(255) NOT NULL,
    missed_total INTEGER DEFAULT 0,
    missed_bad INTEGER DEFAULT 0,
    missed_excused INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(summary_group_id, full_name)
);

CREATE INDEX IF NOT EXISTS idx_groups_department_id ON groups(department_id);
CREATE INDEX IF NOT EXISTS idx_students_group_id ON students(group_id);
CREATE INDEX IF NOT EXISTS idx_attendance_student_id ON attendance(student_id);
CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(date);
CREATE INDEX IF NOT EXISTS idx_specialties_department_id ON specialties(department_id);
CREATE INDEX IF NOT EXISTS idx_summary_groups_specialty_id ON summary_groups(specialty_id);
CREATE INDEX IF NOT EXISTS idx_summary_students_group_id ON summary_students(summary_group_id);
`
}
