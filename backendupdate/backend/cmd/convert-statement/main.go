package main

import (
	"flag"
	"fmt"
	"log"
	"path/filepath"

	"dashboard/internal/config"
	"dashboard/internal/converter"
)

// CLI-утилита для конвертации vedomost.xls -> vedomost.json с использованием config.
func main() {
	inputFlag := flag.String("in", "", "Путь к файлу vedomost.xls/.xlsx (по умолчанию из конфигурации)")
	outputFlag := flag.String("out", "", "Путь к vedomost.json (по умолчанию из конфигурации)")
	flag.Parse()

	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("Ошибка загрузки конфига: %v", err)
	}

	input := *inputFlag
	if input == "" {
		input = cfg.StatementInput
	}

	output := *outputFlag
	if output == "" {
		output = cfg.StatementOutput
	}

	fmt.Printf("Конвертация ведомости\n")
	fmt.Printf("  Вход:  %s\n", input)
	fmt.Printf("  Выход: %s\n", output)

	if !filepath.IsAbs(output) {
		if abs, err := filepath.Abs(output); err == nil {
			output = abs
		}
	}

	if err := converter.ConvertStatement(input, output, cfg.PythonScript); err != nil {
		log.Fatalf("Ошибка конвертации ведомости: %v", err)
	}

	fmt.Println("✅ vedomost.json успешно обновлён")
}

