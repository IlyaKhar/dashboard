package main

import (
	"context"
	"fmt"
	"log"
	"net/http"
	"os"
	"os/signal"
	"path/filepath"
	"syscall"
	"time"

	"dashboard/internal/api"
	"dashboard/internal/config"
	"dashboard/internal/database"
	"dashboard/internal/fetcher"
	"dashboard/internal/middleware"
	"dashboard/internal/scheduler"
	"dashboard/internal/services"
	"dashboard/internal/utils"

	"github.com/gin-gonic/gin"
	"github.com/robfig/cron/v3"
)

func main() {
	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("[Server] Ошибка загрузки конфигурации: %v", err)
	}

	log.Printf("[Server] Запуск бэкенд сервера...")
	log.Printf("[Server] Корневая директория: %s", cfg.ProjectRoot)
	log.Printf("[Server] Интервал обновления: %v", cfg.RefreshInterval)

	gin.SetMode(gin.ReleaseMode)

	if cfg.DatabaseURL != "" {
		if err := database.Connect(cfg.DatabaseURL); err != nil {
			log.Printf("[Server] Предупреждение: не удалось подключиться к БД: %v", err)
			log.Println("[Server] Продолжаем работу без БД (данные не будут сохраняться)")
		} else {
			if err := database.InitSchema(); err != nil {
				log.Printf("[Server] Предупреждение: не удалось инициализировать схему БД: %v", err)
			}
			defer database.Close()
		}
	} else {
		log.Println("[Server] DATABASE_URL не указан, работаем без БД")
	}

	dbLoader := database.NewLoader(database.DB)
	converterDir := filepath.Join(cfg.ProjectRoot, "backend", "internal", "converter")

	sched := scheduler.NewScheduler(
		cfg.ProjectRoot,
		cfg.AttendanceInput,
		cfg.AttendanceOutput,
		cfg.StatementInput,
		cfg.StatementOutput,
		cfg.StudentsInput,
		cfg.StudentsOutput,
		cfg.LessonsInput,
		cfg.LessonsOutput,
		cfg.ScheduleGridInput,
		cfg.ScheduleGridOutput,
		cfg.PythonScript,
	)

	attendanceService := services.NewAttendanceService(cfg.AttendanceOutput, cfg.StudentsOutput, cfg.StatementOutput)
	scheduleService := services.NewScheduleService(cfg.LessonsOutput)
	reconciliationService := services.NewReconciliationService(attendanceService, scheduleService)
	lessonsService := services.NewLessonsService(database.DB)
	dashboardMainService := services.NewDashboardService(database.DB)
	// Настраиваем dashboardMainService для работы с JSON файлами
	dashboardMainService.SetAttendanceService(attendanceService)
	dashboardMainService.SetStudentsPath(cfg.StudentsOutput)
	thresholdsService := services.NewThresholdsService(database.DB)
	// Более строгий лимит логина: 5 попыток за 5 минут с одного IP.
	loginRateLimiter := middleware.NewRateLimiter(5, 5*time.Minute)
	refreshHistory := api.NewRefreshHistoryStore(50)

	router := api.SetupRouter(
		cfg, sched, dbLoader,
		attendanceService, scheduleService, reconciliationService,
		lessonsService, dashboardMainService, thresholdsService,
		loginRateLimiter,
		refreshHistory,
	)

	serverAddr := fmt.Sprintf("%s:%s", cfg.ServerHost, cfg.ServerPort)
	httpServer := &http.Server{Addr: serverAddr, Handler: router}

	go func() {
		log.Printf("[Server] HTTP сервер запущен на http://%s", serverAddr)
		log.Printf("[Server] Swagger UI доступен на http://%s/swagger/", serverAddr)
		if err := httpServer.ListenAndServe(); err != nil && err != http.ErrServerClosed {
			log.Fatalf("[Server] Ошибка запуска HTTP сервера: %v", err)
		}
	}()

	c := cron.New()
	cronExpr := formatCronInterval(cfg.RefreshInterval)
	_, err = c.AddFunc(cronExpr, func() {
		log.Println("[Server] Запуск автоматического обновления данных...")
		utils.SyncFromOneC(cfg.OneCSourceDir, converterDir)
		if err := sched.RefreshData(); err != nil {
			log.Printf("[Server] Ошибка обновления данных: %v", err)
			refreshHistory.AddEvent("error", err.Error())
			return
		}
		if database.DB != nil {
			_ = dbLoader.LoadAttendance(cfg.AttendanceOutput)
			_ = dbLoader.LoadStatement(cfg.StatementOutput)
			_ = dbLoader.LoadLessons(cfg.LessonsOutput)
		}
		refreshHistory.AddEvent("success", "Автоматическое обновление выполнено")
	})
	if err != nil {
		log.Fatalf("[Server] Ошибка настройки cron: %v", err)
	}

	// Расписание с сайта ОКЭИ:
	// 1) раз в неделю по cron (понедельник 03:00)
	// 2) один раз при старте сервера, чтобы не ждать cron.
	if cfg.OkseiScheduleURL != "" {
		// Еженедельный cron
		_, err = c.AddFunc("0 3 * * 1", func() {
			log.Println("[Server] Запуск выгрузки расписания с сайта ОКЭИ (cron)...")
			if err := fetcher.FetchScheduleFromOkseiPage(
				cfg.OkseiScheduleURL,
				cfg.OkseiScheduleOutput,
				cfg.PythonScript,
				converterDir,
			); err != nil {
				log.Printf("[Server] Ошибка выгрузки расписания ОКЭИ (cron): %v", err)
				return
			}
			log.Println("[Server] Расписание с сайта ОКЭИ сохранено в", cfg.OkseiScheduleOutput)
		})
		if err != nil {
			log.Printf("[Server] Предупреждение: не удалось добавить cron расписания ОКЭИ: %v", err)
		} else {
			log.Println("[Server] Cron расписания ОКЭИ: понедельник 03:00")
		}

		// Первичная выгрузка при старте
		log.Println("[Server] Первичная выгрузка расписания ОКЭИ при старте...")
		if err := fetcher.FetchScheduleFromOkseiPage(
			cfg.OkseiScheduleURL,
			cfg.OkseiScheduleOutput,
			cfg.PythonScript,
			converterDir,
		); err != nil {
			log.Printf("[Server] Ошибка первичной выгрузки расписания ОКЭИ: %v", err)
		} else {
			log.Printf("[Server] Расписание ОКЭИ сохранено в %s", cfg.OkseiScheduleOutput)
		}
	}

	log.Println("[Server] Первоначальное обновление данных...")
	utils.SyncFromOneC(cfg.OneCSourceDir, converterDir)
	if err := sched.RefreshData(); err != nil {
		log.Printf("[Server] Предупреждение при первоначальном обновлении: %v", err)
		refreshHistory.AddEvent("error", err.Error())
	} else {
		refreshHistory.AddEvent("success", "Первоначальное обновление выполнено")
	}
	if database.DB != nil {
		_ = dbLoader.LoadAttendance(cfg.AttendanceOutput)
		_ = dbLoader.LoadStatement(cfg.StatementOutput)
		_ = dbLoader.LoadLessons(cfg.LessonsOutput)
	}

	c.Start()
	log.Printf("[Server] Планировщик запущен. Обновление данных каждые %v.", cfg.RefreshInterval)
	log.Println("[Server] Нажмите Ctrl+C для остановки...")

	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, os.Interrupt, syscall.SIGTERM)
	<-sigChan

	log.Println("[Server] Получен сигнал завершения. Остановка сервера...")
	c.Stop()

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := httpServer.Shutdown(ctx); err != nil {
		log.Printf("[Server] Ошибка остановки HTTP сервера: %v", err)
	}
	log.Println("[Server] Сервер остановлен.")
}

func formatCronInterval(d time.Duration) string {
	minutes := int(d.Minutes())
	if minutes < 60 {
		return fmt.Sprintf("@every %dm", minutes)
	}
	hours := minutes / 60
	if hours*60 == minutes {
		return fmt.Sprintf("@every %dh", hours)
	}
	return fmt.Sprintf("@every %dm", minutes)
}
