package fetcher

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"dashboard/internal/converter"
)

const (
	httpTimeout    = 60 * time.Second
	okseiPageURL   = "https://oksei.ru/studentu/raspisanie_uchebnykh_zanyatij"
	xlsLinkPattern = `(?i)href\s*=\s*["']([^"']+\.xls(?:\?[^"']*)?)["']`
)

// FetchSchedule загружает расписание по прямой ссылке на файл (.xlsx/.xls) и сохраняет в JSON.
func FetchSchedule(fileURL string, outputPath string) error {
	if fileURL == "" {
		return nil
	}
	log.Printf("[Fetcher] Загрузка расписания с %s", fileURL)
	client := &http.Client{Timeout: httpTimeout}
	resp, err := client.Get(fileURL)
	if err != nil {
		return fmt.Errorf("ошибка загрузки: %v", err)
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("сервер вернул статус %d", resp.StatusCode)
	}
	contentType := resp.Header.Get("Content-Type")
	disp := resp.Header.Get("Content-Disposition")
	isExcel := strings.Contains(contentType, "spreadsheet") ||
		strings.Contains(contentType, "excel") ||
		strings.Contains(contentType, "vnd.openxmlformats") ||
		strings.HasSuffix(strings.ToLower(fileURL), ".xlsx") ||
		strings.HasSuffix(strings.ToLower(fileURL), ".xls") ||
		strings.Contains(strings.ToLower(disp), ".xlsx") ||
		strings.Contains(strings.ToLower(disp), ".xls")
	if !isExcel {
		return fmt.Errorf("поддерживается только Excel (.xlsx/.xls). Content-Type: %s", contentType)
	}
	return fetchAndConvertExcel(resp.Body, fileURL, outputPath, "")
}

// FetchScheduleFromOkseiPage загружает страницу расписания ОКЭИ, находит ссылки на .xls,
// скачивает файлы, парсит timetable (группа → день → пара → дисциплина) и сохраняет в outputPath.
// Для .xls использует Python-скрипт из pythonScriptPath для конвертации в .xlsx.
// Если converterDir не пустой, первый найденный .xls дополнительно сохраняется в эту директорию
// под именем "расписание.xls" для дальнейшей обработки конвертерами (сетка расписания).
func FetchScheduleFromOkseiPage(pageURL, outputPath, pythonScriptPath, converterDir string) error {
	if pageURL == "" {
		pageURL = okseiPageURL
	}
	log.Printf("[Fetcher] Загрузка страницы расписания: %s", pageURL)
	client := &http.Client{Timeout: httpTimeout}
	resp, err := client.Get(pageURL)
	if err != nil {
		return fmt.Errorf("ошибка загрузки страницы: %v", err)
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("страница вернула статус %d", resp.StatusCode)
	}
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return fmt.Errorf("ошибка чтения страницы: %v", err)
	}
	base, _ := url.Parse(pageURL)
	links := extractXLSLinks(body, base)
	if len(links) == 0 {
		return fmt.Errorf("на странице не найдено ссылок на .xls файлы")
	}
	log.Printf("[Fetcher] Найдено ссылок на .xls: %d", len(links))
	tmpDir := filepath.Join(os.TempDir(), "oksei-schedule")
	if err := os.MkdirAll(tmpDir, 0o755); err != nil {
		return fmt.Errorf("временная папка: %v", err)
	}
	var outputs []*converter.OkseiTimetableOutput
	for i, link := range links {
		xlsPath := filepath.Join(tmpDir, fmt.Sprintf("schedule_%d_%d.xls", time.Now().Unix(), i))
		if err := downloadFile(client, link, xlsPath); err != nil {
			log.Printf("[Fetcher] Пропуск файла %s: %v", link, err)
			continue
		}

		// Дополнительно сохраняем первый файл расписания в директорию конвертера,
		// чтобы его могли обработать ConvertScheduleGrid/ConvertScheduleGridToLessonsFormat.
		if converterDir != "" && i == 0 {
			if err := os.MkdirAll(converterDir, 0o755); err != nil {
				log.Printf("[Fetcher] Не удалось создать директорию конвертера для расписания: %v", err)
			} else {
				fixedPath := filepath.Join(converterDir, "расписание.xls")
				if err := copyFileLocal(xlsPath, fixedPath); err != nil {
					log.Printf("[Fetcher] Не удалось сохранить файл расписания в конвертер: %v", err)
				} else {
					log.Printf("[Fetcher] Файл расписания сохранён в конвертер: %s", fixedPath)
				}
			}
		}

		// Конвертируем .xls → .xlsx, затем парсим через excelize внутри ParseOkseiTimetable.
		xlsxPath := strings.TrimSuffix(xlsPath, ".xls") + ".xlsx"
		if err := converter.ConvertXLSToXLSX(xlsPath, xlsxPath, pythonScriptPath); err != nil {
			log.Printf("[Fetcher] Ошибка конвертации %s: %v", link, err)
			_ = os.Remove(xlsPath)
			_ = os.Remove(xlsxPath)
			continue
		}
		_ = os.Remove(xlsPath)

		out, err := converter.ParseOkseiTimetable(xlsxPath)
		_ = os.Remove(xlsxPath)
		if err != nil {
			log.Printf("[Fetcher] Парсинг %s: %v", link, err)
			continue
		}
		outputs = append(outputs, out)
	}
	if len(outputs) == 0 {
		return fmt.Errorf("ни один .xls файл не удалось обработать")
	}
	merged := mergeOkseiTimetables(outputs)
	jsonData, err := json.MarshalIndent(merged, "", "  ")
	if err != nil {
		return fmt.Errorf("сериализация JSON: %v", err)
	}
	if err := os.MkdirAll(filepath.Dir(outputPath), 0o755); err != nil {
		return fmt.Errorf("создание папки: %v", err)
	}
	if err := os.WriteFile(outputPath, jsonData, 0o644); err != nil {
		return fmt.Errorf("запись файла: %v", err)
	}
	log.Printf("[Fetcher] Расписание сохранено в %s (объединено файлов: %d)", outputPath, len(outputs))
	return nil
}

func mergeOkseiTimetables(outputs []*converter.OkseiTimetableOutput) *converter.OkseiTimetableOutput {
	merged := outputs[0]
	for i := 1; i < len(outputs); i++ {
		o := outputs[i]
		merged.Period = merged.Period + " | " + o.Period
		byGroup := make(map[string]*converter.OkseiGroupSchedule)
		for j := range merged.Groups {
			g := &merged.Groups[j]
			byGroup[g.Group] = g
		}
		for j := range o.Groups {
			g := &o.Groups[j]
			if existing, ok := byGroup[g.Group]; ok {
				for day, pairs := range g.Schedule {
					if existing.Schedule[day] == nil {
						existing.Schedule[day] = make(map[int]converter.OkseiLesson)
					}
					for p, l := range pairs {
						existing.Schedule[day][p] = l
					}
				}
			} else {
				merged.Groups = append(merged.Groups, *g)
				byGroup[g.Group] = &merged.Groups[len(merged.Groups)-1]
			}
		}
	}
	return merged
}

func extractXLSLinks(html []byte, base *url.URL) []string {
	re := regexp.MustCompile(xlsLinkPattern)
	matches := re.FindAllSubmatch(html, -1)
	var out []string
	seen := make(map[string]bool)
	for _, m := range matches {
		if len(m) < 2 {
			continue
		}
		raw := strings.TrimSpace(string(m[1]))
		href, _ := url.Parse(raw)
		if href == nil {
			continue
		}
		resolved := base.ResolveReference(href)
		u := resolved.String()
		if seen[u] {
			continue
		}
		seen[u] = true
		out = append(out, u)
	}
	return out
}

func downloadFile(client *http.Client, fileURL, destPath string) error {
	resp, err := client.Get(fileURL)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("статус %d", resp.StatusCode)
	}
	f, err := os.Create(destPath)
	if err != nil {
		return err
	}
	_, err = io.Copy(f, resp.Body)
	_ = f.Close()
	if err != nil {
		_ = os.Remove(destPath)
		return err
	}
	return nil
}

// copyFileLocal копирует файл src → dst (простая обёртка для локального копирования).
func copyFileLocal(src, dst string) error {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer in.Close()

	out, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer func() {
		_ = out.Close()
	}()

	if _, err := io.Copy(out, in); err != nil {
		_ = os.Remove(dst)
		return err
	}

	return nil
}

// fetchAndConvertExcel скачивает Excel в файл, при необходимости конвертирует .xls→.xlsx, затем в JSON.
func fetchAndConvertExcel(body io.Reader, fileURL, outputPath, pythonScriptPath string) error {
	tmpDir := filepath.Join(os.TempDir(), "oksei-schedule")
	if err := os.MkdirAll(tmpDir, 0o755); err != nil {
		return fmt.Errorf("ошибка создания временной папки: %v", err)
	}
	ext := ".xlsx"
	lowerURL := strings.ToLower(fileURL)
	if strings.HasSuffix(lowerURL, ".xls") && !strings.HasSuffix(lowerURL, ".xlsx") {
		ext = ".xls"
	}
	tmpFile := filepath.Join(tmpDir, fmt.Sprintf("schedule_%d%s", time.Now().Unix(), ext))
	out, err := os.Create(tmpFile)
	if err != nil {
		return fmt.Errorf("ошибка создания временного файла: %v", err)
	}
	written, err := io.Copy(out, body)
	_ = out.Close()
	if err != nil {
		_ = os.Remove(tmpFile)
		return fmt.Errorf("ошибка сохранения файла: %v", err)
	}
	log.Printf("[Fetcher] Скачано %d байт", written)
	xlsxPath := tmpFile
	if ext == ".xls" && pythonScriptPath != "" {
		xlsxPath = strings.TrimSuffix(tmpFile, ".xls") + ".xlsx"
		if err := converter.ConvertXLSToXLSX(tmpFile, xlsxPath, pythonScriptPath); err != nil {
			_ = os.Remove(tmpFile)
			return fmt.Errorf("ошибка конвертации XLS→XLSX: %v", err)
		}
		_ = os.Remove(tmpFile)
	}
	if err := converter.ConvertLessons(xlsxPath, outputPath); err != nil {
		_ = os.Remove(xlsxPath)
		return fmt.Errorf("ошибка конвертации в JSON: %v", err)
	}
	_ = os.Remove(xlsxPath)
	log.Printf("[Fetcher] Расписание сохранено в %s", outputPath)
	return nil
}

