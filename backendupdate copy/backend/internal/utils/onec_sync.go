package utils

import (
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
)

// SyncFromOneC копирует Excel-файлы из директории 1С (\\1C01\proba)
// в директорию конвертеров (backend/internal/converter).
// Если каких-то файлов нет — просто пишем предупреждение и идём дальше.
func SyncFromOneC(sourceDir, converterDir string) {
	if sourceDir == "" {
		log.Println("[OneC] ONEC_SOURCE_DIR не задан, пропускаем синхронизацию")
		return
	}

	info, err := os.Stat(sourceDir)
	if err != nil {
		log.Printf("[OneC] Не удалось получить доступ к директории %s: %v", sourceDir, err)
		return
	}
	if !info.IsDir() {
		log.Printf("[OneC] Путь %s не является директорией, пропускаем синхронизацию", sourceDir)
		return
	}

	if _, err := os.Stat(converterDir); err != nil {
		if os.IsNotExist(err) {
			if mkErr := os.MkdirAll(converterDir, 0o755); mkErr != nil {
				log.Printf("[OneC] Не удалось создать директорию конвертеров %s: %v", converterDir, mkErr)
				return
			}
		} else {
			log.Printf("[OneC] Ошибка доступа к директории конвертеров %s: %v", converterDir, err)
			return
		}
	}

	files := []string{
		"Посещаемость.xlsx",
		"ведомость.xls",
		"ведомость.xlsx",
		"Ведомостьколва.xlsx",
		"Проба.xlsx",
	}

	log.Printf("[OneC] Синхронизация файлов из %s в %s...", sourceDir, converterDir)

	for _, name := range files {
		src := filepath.Join(sourceDir, name)
		dst := filepath.Join(converterDir, name)

		if _, err := os.Stat(src); os.IsNotExist(err) {
			log.Printf("[OneC] Файл не найден (пропускаем): %s", src)
			continue
		}

		if err := copyFile(src, dst); err != nil {
			log.Printf("[OneC] Ошибка копирования %s → %s: %v", src, dst, err)
			continue
		}

		log.Printf("[OneC] Обновлён файл: %s", dst)
	}
}

func copyFile(src, dst string) error {
	in, err := os.Open(src)
	if err != nil {
		return fmt.Errorf("ошибка открытия исходного файла: %w", err)
	}
	defer in.Close()

	out, err := os.Create(dst)
	if err != nil {
		return fmt.Errorf("ошибка создания целевого файла: %w", err)
	}
	defer func() {
		_ = out.Close()
	}()

	if _, err := io.Copy(out, in); err != nil {
		return fmt.Errorf("ошибка копирования данных: %w", err)
	}

	if err := out.Sync(); err != nil {
		return fmt.Errorf("ошибка синхронизации файла: %w", err)
	}

	return nil
}

