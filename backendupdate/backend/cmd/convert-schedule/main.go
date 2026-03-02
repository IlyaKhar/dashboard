package main

import (
	"flag"
	"fmt"
	"log"
	"path/filepath"

	"dashboard/internal/config"
	"dashboard/internal/converter"
)

// CLI-утилита для конвертации расписания Excel -> schedule.json с использованием config.
func main() {
	inputFlag := flag.String("in", "", "Путь к файлу расписания (.xlsx) (по умолчанию из конфигурации)")
	outputFlag := flag.String("out", "", "Путь к schedule.json (по умолчанию из конфигурации)")
	flag.Parse()

	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("Ошибка загрузки конфига: %v", err)
	}

	input := *inputFlag
	if input == "" {
		input = cfg.LessonsInput
	}

	output := *outputFlag
	if output == "" {
		output = cfg.LessonsOutput
	}

	fmt.Printf("Конвертация расписания\n")
	fmt.Printf("  Вход:  %s\n", input)
	fmt.Printf("  Выход: %s\n", output)

	if !filepath.IsAbs(output) {
		if abs, err := filepath.Abs(output); err == nil {
			output = abs
		}
	}

	if err := converter.ConvertLessons(input, output); err != nil {
		log.Fatalf("Ошибка конвертации расписания: %v", err)
	}

	fmt.Println("✅ schedule.json успешно обновлён")
}

