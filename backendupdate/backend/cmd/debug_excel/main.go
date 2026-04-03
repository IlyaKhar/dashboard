package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	projectRoot := "/Users/maloy/Desktop/dashboard_OKSEI-master"
	
	possiblePaths := []string{
		filepath.Join(projectRoot, "backend", "internal", "converter", "Проба.xlsx"),
		filepath.Join(projectRoot, "Проба.xlsx"),
	}
	
	var inputFile string
	for _, path := range possiblePaths {
		if _, err := os.Stat(path); err == nil {
			inputFile = path
			fmt.Printf("Найден файл: %s\n\n", path)
			break
		}
	}
	
	if inputFile == "" {
		fmt.Fprintf(os.Stderr, "Файл Проба.xlsx не найден\n")
		os.Exit(1)
	}
	
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Ошибка открытия: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()
	
	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Ошибка чтения: %v\n", err)
		os.Exit(1)
	}
	
	fmt.Printf("Всего строк: %d\n\n", len(rows))
	
	// Выводим первые 50 строк с данными
	for i := 0; i < len(rows) && i < 50; i++ {
		row := rows[i]
		if len(row) == 0 {
			continue
		}
		
		fmt.Printf("=== Строка %d (колонок: %d) ===\n", i+1, len(row))
		for j := 0; j < len(row) && j < 10; j++ {
			val := strings.TrimSpace(row[j])
			if len(val) > 40 {
				val = val[:40] + "..."
			}
			fmt.Printf("  [%d] '%s'\n", j, val)
		}
		fmt.Println()
	}
}
