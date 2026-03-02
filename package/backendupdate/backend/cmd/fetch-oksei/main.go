package main

import (
	"log"

	"dashboard/internal/config"
	"dashboard/internal/fetcher"
)

// Небольшая CLI-утилита, которая один раз тянет расписание с сайта ОКЭИ
// и сохраняет его в JSON (по путям из config: OkseiScheduleURL, OkseiScheduleOutput).
//
// Пример запуска:
//
//	cd backend
//	go run ./cmd/fetch-oksei
//
// На сервере можно собрать бинарник:
//
//	go build -o fetch-oksei ./cmd/fetch-oksei
//
// и в cron/systemd дергать этот бинарник по расписанию.
func main() {
	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("[fetch-oksei] Ошибка загрузки конфига: %v", err)
	}

	if cfg.OkseiScheduleURL == "" {
		log.Fatalf("[fetch-oksei] Не задан URL расписания ОКЭИ (OKSEI_SCHEDULE_URL)")
	}

	log.Printf("[fetch-oksei] Загрузка расписания с %s", cfg.OkseiScheduleURL)

	if err := fetcher.FetchScheduleFromOkseiPage(cfg.OkseiScheduleURL, cfg.OkseiScheduleOutput, cfg.PythonScript); err != nil {
		log.Fatalf("[fetch-oksei] Ошибка выгрузки расписания: %v", err)
	}

	log.Printf("[fetch-oksei] Расписание сохранено в %s", cfg.OkseiScheduleOutput)
}
