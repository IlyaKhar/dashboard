package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"dashboard/internal/converter"
)

func main() {
	// Определяем пути к файлам
	projectRoot := "/Users/maloy/Desktop/dashboard_OKSEI-master"
	
	// Пробуем разные возможные пути к файлу
	possiblePaths := []string{
		filepath.Join(projectRoot, "backend", "internal", "converter", "Проба.xlsx"),
		filepath.Join(projectRoot, "Проба.xlsx"),
	}
	
	var inputFile string
	for _, path := range possiblePaths {
		if _, err := os.Stat(path); err == nil {
			inputFile = path
			fmt.Printf("Найден файл: %s\n", path)
			break
		}
	}
	
	if inputFile == "" {
		log.Fatal("Файл Проба.xlsx не найден. Проверьте путь.")
	}
	
	outputFile := filepath.Join(projectRoot, "public", "lessons.json")
	
	fmt.Printf("Тестирование конвертера lessons.go\n")
	fmt.Printf("Входной файл: %s\n", inputFile)
	fmt.Printf("Выходной файл: %s\n", outputFile)
	fmt.Println()
	
	// Запускаем конвертацию
	if err := converter.ConvertLessons(inputFile, outputFile); err != nil {
		log.Fatalf("Ошибка конвертации: %v", err)
	}
	
	fmt.Println()
	fmt.Println("✅ Конвертация успешно завершена!")
	fmt.Printf("Результат сохранён в: %s\n", outputFile)
}
