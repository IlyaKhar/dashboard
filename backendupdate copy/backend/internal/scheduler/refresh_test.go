package scheduler

import (
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestScheduler_shouldUpdateFile(t *testing.T) {
	// Создаём временную директорию для тестов
	tmpDir := t.TempDir()

	inputFile := filepath.Join(tmpDir, "input.txt")
	outputFile := filepath.Join(tmpDir, "output.txt")

	s := &Scheduler{
		lastModified: make(map[string]time.Time),
	}

	// Тест 1: Входной файл не существует
	shouldUpdate, err := s.shouldUpdateFile(inputFile, outputFile)
	if err == nil {
		t.Error("Ожидалась ошибка для несуществующего входного файла")
	}

	// Создаём входной файл
	if err := os.WriteFile(inputFile, []byte("test"), 0644); err != nil {
		t.Fatalf("Ошибка создания входного файла: %v", err)
	}

	// Тест 2: Выходной файл не существует - должно вернуть true
	shouldUpdate, err = s.shouldUpdateFile(inputFile, outputFile)
	if err != nil {
		t.Fatalf("Неожиданная ошибка: %v", err)
	}
	if !shouldUpdate {
		t.Error("Ожидалось shouldUpdate=true, когда выходной файл не существует")
	}

	// Создаём выходной файл
	if err := os.WriteFile(outputFile, []byte("test"), 0644); err != nil {
		t.Fatalf("Ошибка создания выходного файла: %v", err)
	}

	// Ждём немного, чтобы время модификации было разным
	time.Sleep(10 * time.Millisecond)

	// Тест 3: Входной файл новее выходного - должно вернуть true
	if err := os.WriteFile(inputFile, []byte("updated"), 0644); err != nil {
		t.Fatalf("Ошибка обновления входного файла: %v", err)
	}

	shouldUpdate, err = s.shouldUpdateFile(inputFile, outputFile)
	if err != nil {
		t.Fatalf("Неожиданная ошибка: %v", err)
	}
	if !shouldUpdate {
		t.Error("Ожидалось shouldUpdate=true, когда входной файл новее")
	}

	// Обновляем выходной файл
	if err := os.WriteFile(outputFile, []byte("updated"), 0644); err != nil {
		t.Fatalf("Ошибка обновления выходного файла: %v", err)
	}

	// Тест 4: Выходной файл новее входного - должно вернуть false
	shouldUpdate, err = s.shouldUpdateFile(inputFile, outputFile)
	if err != nil {
		t.Fatalf("Неожиданная ошибка: %v", err)
	}
	if shouldUpdate {
		t.Error("Ожидалось shouldUpdate=false, когда выходной файл новее")
	}
}
