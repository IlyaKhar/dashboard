package main

import (
	"flag"
	"fmt"
	"log"
	"path/filepath"

	"dashboard/internal/config"
	"dashboard/internal/converter"
)

// CLI-утилита для мастер-конвертации: один файл ведомость.xls → все JSON
func main() {
	inputFlag := flag.String("in", "", "Путь к файлу ведомость.xls/.xlsx (по умолчанию из конфигурации)")
	outputDirFlag := flag.String("out", "", "Директория для сохранения JSON (по умолчанию public/)")
	flag.Parse()

	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("Ошибка загрузки конфига: %v", err)
	}

	input := *inputFlag
	if input == "" {
		input = cfg.StatementInput
	}

	outputDir := *outputDirFlag
	if outputDir == "" {
		outputDir = filepath.Join(cfg.ProjectRoot, "public")
	}

	fmt.Printf("Мастер-конвертация ведомости\n")
	fmt.Printf("  Вход:  %s\n", input)
	fmt.Printf("  Выход: %s\n", outputDir)

	if !filepath.IsAbs(outputDir) {
		if abs, err := filepath.Abs(outputDir); err == nil {
			outputDir = abs
		}
	}

	result, err := converter.ConvertMaster(input, outputDir, cfg.PythonScript)
	if err != nil {
		log.Fatalf("Ошибка мастер-конвертации: %v", err)
	}

	fmt.Println("\n✅ Мастер-конвертация завершена!")
	if result.StudentsOutput != "" {
		fmt.Printf("  ✓ students.json: %s\n", result.StudentsOutput)
	}
	if result.AttendanceOutput != "" {
		fmt.Printf("  ✓ attendance.json: %s\n", result.AttendanceOutput)
	}
	if result.VedomostOutput != "" {
		fmt.Printf("  ✓ vedomost.json: %s\n", result.VedomostOutput)
	}

	if len(result.Warnings) > 0 {
		fmt.Println("\n⚠️  Предупреждения:")
		for _, w := range result.Warnings {
			fmt.Printf("  - %s\n", w)
		}
	}

	if len(result.Errors) > 0 {
		fmt.Println("\n❌ Ошибки:")
		for _, e := range result.Errors {
			fmt.Printf("  - %s\n", e)
		}
	}
}
