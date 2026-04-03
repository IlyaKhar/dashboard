-- Схема базы данных для дашборда посещаемости

-- Таблица отделений
CREATE TABLE IF NOT EXISTS departments (
    id SERIAL PRIMARY KEY,
    name VARCHAR(255) NOT NULL UNIQUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Таблица групп
CREATE TABLE IF NOT EXISTS groups (
    id SERIAL PRIMARY KEY,
    department_id INTEGER NOT NULL REFERENCES departments(id) ON DELETE CASCADE,
    name VARCHAR(50) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(department_id, name)
);

-- Таблица студентов
CREATE TABLE IF NOT EXISTS students (
    id SERIAL PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES groups(id) ON DELETE CASCADE,
    full_name VARCHAR(255) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(group_id, full_name)
);

-- Таблица посещаемости (attendance)
CREATE TABLE IF NOT EXISTS attendance (
    id SERIAL PRIMARY KEY,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    date DATE NOT NULL,
    missed_hours INTEGER NOT NULL DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(student_id, date)
);

-- Таблица специальностей (для summary)
CREATE TABLE IF NOT EXISTS specialties (
    id SERIAL PRIMARY KEY,
    department_id INTEGER NOT NULL REFERENCES departments(id) ON DELETE CASCADE,
    name VARCHAR(255) NOT NULL,
    total_missed INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(department_id, name)
);

-- Таблица групп в summary (связь с specialty)
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

-- Таблица занятий (расписание по парам)
CREATE TABLE IF NOT EXISTS lessons (
    id SERIAL PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES groups(id) ON DELETE CASCADE,
    date_time TIMESTAMP NOT NULL,
    discipline VARCHAR(255) NOT NULL,
    teacher VARCHAR(255),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(group_id, date_time, discipline)
);

-- Таблица посещаемости по занятиям
CREATE TABLE IF NOT EXISTS lesson_attendance (
    id SERIAL PRIMARY KEY,
    lesson_id INTEGER NOT NULL REFERENCES lessons(id) ON DELETE CASCADE,
    student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    attendance BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(lesson_id, student_id)
);

-- Таблица расписания (связь пар с группами по дням недели)
CREATE TABLE IF NOT EXISTS schedule (
    id SERIAL PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES groups(id) ON DELETE CASCADE,
    day_of_week INTEGER NOT NULL CHECK (day_of_week >= 0 AND day_of_week <= 6), -- 0 = воскресенье, 1 = понедельник, ..., 6 = суббота
    lesson_number INTEGER NOT NULL CHECK (lesson_number >= 1 AND lesson_number <= 8), -- Номер пары (1-8)
    discipline VARCHAR(255) NOT NULL,
    teacher VARCHAR(255),
    start_time TIME NOT NULL, -- Время начала пары (например, 08:30)
    end_time TIME NOT NULL,   -- Время окончания пары (например, 10:00)
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(group_id, day_of_week, lesson_number)
);

-- Таблица порогов цветовой индикации
CREATE TABLE IF NOT EXISTS thresholds (
    id SERIAL PRIMARY KEY,
    type VARCHAR(50) NOT NULL DEFAULT 'lesson', -- Тип: 'lesson' (пара), 'group' (группа), 'department' (отделение)
    green_threshold DECIMAL(5,2) NOT NULL DEFAULT 90.0,  -- Верхний порог для зелёного (>=)
    yellow_threshold DECIMAL(5,2) NOT NULL DEFAULT 70.0, -- Нижний порог для жёлтого (>=)
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(type)
);

-- Индексы для ускорения запросов
CREATE INDEX IF NOT EXISTS idx_groups_department_id ON groups(department_id);
CREATE INDEX IF NOT EXISTS idx_students_group_id ON students(group_id);
CREATE INDEX IF NOT EXISTS idx_attendance_student_id ON attendance(student_id);
CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(date);
CREATE INDEX IF NOT EXISTS idx_specialties_department_id ON specialties(department_id);
CREATE INDEX IF NOT EXISTS idx_summary_groups_specialty_id ON summary_groups(specialty_id);
CREATE INDEX IF NOT EXISTS idx_summary_students_group_id ON summary_students(summary_group_id);
CREATE INDEX IF NOT EXISTS idx_lessons_group_id ON lessons(group_id);
CREATE INDEX IF NOT EXISTS idx_lessons_date_time ON lessons(date_time);
CREATE INDEX IF NOT EXISTS idx_lesson_attendance_lesson_id ON lesson_attendance(lesson_id);
CREATE INDEX IF NOT EXISTS idx_lesson_attendance_student_id ON lesson_attendance(student_id);
CREATE INDEX IF NOT EXISTS idx_schedule_group_id ON schedule(group_id);
CREATE INDEX IF NOT EXISTS idx_schedule_day_of_week ON schedule(day_of_week);

-- Функция для обновления updated_at
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ language 'plpgsql';

-- Триггеры для автоматического обновления updated_at
-- Удаляем существующие триггеры перед созданием (если они есть)
DROP TRIGGER IF EXISTS update_departments_updated_at ON departments;
DROP TRIGGER IF EXISTS update_groups_updated_at ON groups;
DROP TRIGGER IF EXISTS update_students_updated_at ON students;
DROP TRIGGER IF EXISTS update_attendance_updated_at ON attendance;
DROP TRIGGER IF EXISTS update_specialties_updated_at ON specialties;
DROP TRIGGER IF EXISTS update_summary_groups_updated_at ON summary_groups;
DROP TRIGGER IF EXISTS update_summary_students_updated_at ON summary_students;
DROP TRIGGER IF EXISTS update_lessons_updated_at ON lessons;
DROP TRIGGER IF EXISTS update_lesson_attendance_updated_at ON lesson_attendance;
DROP TRIGGER IF EXISTS update_schedule_updated_at ON schedule;
DROP TRIGGER IF EXISTS update_thresholds_updated_at ON thresholds;

CREATE TRIGGER update_departments_updated_at BEFORE UPDATE ON departments
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_groups_updated_at BEFORE UPDATE ON groups
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_students_updated_at BEFORE UPDATE ON students
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_attendance_updated_at BEFORE UPDATE ON attendance
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_specialties_updated_at BEFORE UPDATE ON specialties
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_summary_groups_updated_at BEFORE UPDATE ON summary_groups
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_summary_students_updated_at BEFORE UPDATE ON summary_students
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_lessons_updated_at BEFORE UPDATE ON lessons
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_lesson_attendance_updated_at BEFORE UPDATE ON lesson_attendance
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_schedule_updated_at BEFORE UPDATE ON schedule
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_thresholds_updated_at BEFORE UPDATE ON thresholds
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- Вставляем дефолтные пороги для занятий (если их ещё нет)
INSERT INTO thresholds (type, green_threshold, yellow_threshold)
VALUES ('lesson', 90.0, 70.0)
ON CONFLICT (type) DO NOTHING;
