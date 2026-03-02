package api

import (
	"log"
	"net/http"
	"os"
	"path/filepath"
	"sync"
	"time"

	"github.com/gin-gonic/gin"

	"dashboard/internal/config"
	"dashboard/internal/converter"
	"dashboard/internal/database"
	"dashboard/internal/scheduler"
)

// RefreshHistoryItem элемент истории обновлений
type RefreshHistoryItem struct {
	Time    string `json:"time"`
	Status  string `json:"status"`
	Message string `json:"message"`
}

// RefreshHistoryStore хранит историю обновлений (in-memory)
type RefreshHistoryStore struct {
	mu     sync.Mutex
	events []RefreshHistoryItem
	maxLen int
}

// NewRefreshHistoryStore создаёт хранилище истории (хранит последние maxLen событий)
func NewRefreshHistoryStore(maxLen int) *RefreshHistoryStore {
	if maxLen <= 0 {
		maxLen = 50
	}
	return &RefreshHistoryStore{events: nil, maxLen: maxLen}
}

// AddEvent добавляет событие в историю
func (s *RefreshHistoryStore) AddEvent(status, message string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	item := RefreshHistoryItem{
		Time:    time.Now().Format(time.RFC3339),
		Status:  status,
		Message: message,
	}
	s.events = append(s.events, item)
	if len(s.events) > s.maxLen {
		s.events = s.events[len(s.events)-s.maxLen:]
	}
}

// GetEvents возвращает копию истории
func (s *RefreshHistoryStore) GetEvents() []RefreshHistoryItem {
	s.mu.Lock()
	defer s.mu.Unlock()
	if len(s.events) == 0 {
		return nil
	}
	out := make([]RefreshHistoryItem, len(s.events))
	copy(out, s.events)
	return out
}

// GinHandler содержит обработчики API для Gin
type GinHandler struct {
	scheduler         *scheduler.Scheduler
	dbLoader          *database.Loader
	refreshHistory    *RefreshHistoryStore
	lastRefresh       time.Time
	refreshInProgress bool
	mu                sync.Mutex
	cfg               *config.Config
}

func NewGinHandler(scheduler *scheduler.Scheduler, dbLoader *database.Loader, refreshHistory *RefreshHistoryStore) *GinHandler {
	if refreshHistory == nil {
		refreshHistory = NewRefreshHistoryStore(50)
	}
	return &GinHandler{
		scheduler:      scheduler,
		dbLoader:       dbLoader,
		refreshHistory: refreshHistory,
	}
}

// RefreshData запускает ручное обновление данных
// @Summary Ручное обновление данных
// @Description Запускает конвертацию Excel файлов в JSON и загрузку в БД
// @Tags admin
// @Accept json
// @Produce json
// @Success 200 {object} map[string]interface{} "Данные успешно обновлены"
// @Failure 409 {object} map[string]string "Обновление уже выполняется"
// @Failure 500 {object} map[string]string "Ошибка обновления данных"
// @Router /admin/refresh-data [post]
func (h *GinHandler) RefreshData(c *gin.Context) {
	h.mu.Lock()
	if h.refreshInProgress {
		h.mu.Unlock()
		c.JSON(http.StatusConflict, gin.H{
			"error": "Обновление уже выполняется",
		})
		return
	}
	h.refreshInProgress = true
	h.mu.Unlock()

	defer func() {
		h.mu.Lock()
		h.refreshInProgress = false
		h.mu.Unlock()
	}()

	log.Println("[API] Запуск ручного обновления данных...")

	// Запускаем обновление
	if err := h.scheduler.RefreshData(); err != nil {
		log.Printf("[API] Ошибка обновления данных: %v", err)
		c.JSON(http.StatusInternalServerError, gin.H{
			"error":   "Ошибка обновления данных",
			"details": err.Error(),
		})
		return
	}

	// Загружаем в БД (пути из конфига)
	attendancePath := c.GetString("attendance_output")
	statementPath := c.GetString("statement_output")
	
	if attendancePath != "" {
		if err := h.dbLoader.LoadAttendance(attendancePath); err != nil {
			log.Printf("[API] Предупреждение при загрузке посещаемости в БД: %v", err)
		}
	}
	if statementPath != "" {
		if err := h.dbLoader.LoadStatement(statementPath); err != nil {
			log.Printf("[API] Предупреждение при загрузке ведомости в БД: %v", err)
		}
	}

	h.mu.Lock()
	h.lastRefresh = time.Now()
	h.mu.Unlock()

	h.refreshHistory.AddEvent("success", "Данные успешно обновлены")

	c.JSON(http.StatusOK, gin.H{
		"status":  "success",
		"message": "Данные успешно обновлены",
		"time":    h.lastRefresh.Format(time.RFC3339),
	})
}

// GetRefreshStatus возвращает статус последнего обновления
// @Summary Статус обновления данных
// @Description Возвращает информацию о последнем обновлении данных
// @Tags admin
// @Produce json
// @Success 200 {object} map[string]interface{} "Статус обновления"
// @Router /admin/refresh-status [get]
func (h *GinHandler) GetRefreshStatus(c *gin.Context) {
	h.mu.Lock()
	inProgress := h.refreshInProgress
	lastRefresh := h.lastRefresh
	h.mu.Unlock()

	status := gin.H{
		"in_progress": inProgress,
	}

	if !lastRefresh.IsZero() {
		status["last_refresh"] = lastRefresh.Format(time.RFC3339)
		status["last_refresh_ago"] = time.Since(lastRefresh).String()
	} else {
		status["last_refresh"] = nil
		status["last_refresh_ago"] = nil
	}

	c.JSON(http.StatusOK, status)
}

// GetRefreshHistory возвращает историю обновлений (ручные + cron)
// @Summary История обновлений
// @Description Возвращает историю обновлений данных
// @Tags admin
// @Produce json
// @Success 200 {object} map[string]interface{} "История обновлений"
// @Router /admin/refresh-history [get]
func (h *GinHandler) GetRefreshHistory(c *gin.Context) {
	events := h.refreshHistory.GetEvents()
	history := make([]gin.H, 0, len(events))
	for _, e := range events {
		history = append(history, gin.H{
			"time":    e.Time,
			"status":  e.Status,
			"message": e.Message,
		})
	}
	c.JSON(http.StatusOK, gin.H{"history": history})
}

// HealthCheck проверяет работоспособность сервера
// @Summary Health Check
// @Description Проверяет работоспособность сервера
// @Tags system
// @Produce json
// @Success 200 {object} map[string]string "Сервер работает"
// @Router /health [get]
func (h *GinHandler) HealthCheck(c *gin.Context) {
	c.JSON(http.StatusOK, gin.H{
		"status":  "ok",
		"service": "dashboard-backend",
	})
}

// ConvertStatement принимает Excel-файл ведомости и конвертирует его в vedomost.json.
// POST /api/admin/convert/statement (multipart/form-data, поле "file")
func (h *GinHandler) ConvertStatement(c *gin.Context) {
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "field 'file' is required"})
		return
	}

	cfg, err := config.Load()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot load config"})
		return
	}

	tmpDir := filepath.Join(cfg.ProjectRoot, "tmp")
	if err := os.MkdirAll(tmpDir, 0o755); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot create tmp dir"})
		return
	}

	tmpPath := filepath.Join(tmpDir, file.Filename)
	if err := c.SaveUploadedFile(file, tmpPath); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot save uploaded file"})
		return
	}

	if err := converter.ConvertStatement(tmpPath, cfg.StatementOutput, cfg.PythonScript); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	h.refreshHistory.AddEvent("success", "Ведомость сконвертирована через HTTP")

	c.JSON(http.StatusOK, gin.H{
		"ok":      true,
		"message": "vedomost.json обновлён",
		"output":  cfg.StatementOutput,
	})
}

// ConvertSchedule принимает Excel-файл расписания и конвертирует его в schedule.json.
// POST /api/admin/convert/schedule (multipart/form-data, поле "file")
func (h *GinHandler) ConvertSchedule(c *gin.Context) {
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "field 'file' is required"})
		return
	}

	cfg, err := config.Load()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot load config"})
		return
	}

	tmpDir := filepath.Join(cfg.ProjectRoot, "tmp")
	if err := os.MkdirAll(tmpDir, 0o755); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot create tmp dir"})
		return
	}

	tmpPath := filepath.Join(tmpDir, file.Filename)
	if err := c.SaveUploadedFile(file, tmpPath); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot save uploaded file"})
		return
	}

	if err := converter.ConvertLessons(tmpPath, cfg.LessonsOutput); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	h.refreshHistory.AddEvent("success", "Расписание сконвертировано через HTTP")

	c.JSON(http.StatusOK, gin.H{
		"ok":      true,
		"message": "schedule.json обновлён",
		"output":  cfg.LessonsOutput,
	})
}

// ConvertMaster принимает один файл ведомость.xls и генерирует все JSON (students, attendance, vedomost).
// POST /api/admin/convert/master (multipart/form-data, поле "file")
func (h *GinHandler) ConvertMaster(c *gin.Context) {
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "field 'file' is required"})
		return
	}

	cfg, err := config.Load()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot load config"})
		return
	}

	tmpDir := filepath.Join(cfg.ProjectRoot, "tmp")
	if err := os.MkdirAll(tmpDir, 0o755); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot create tmp dir"})
		return
	}

	tmpPath := filepath.Join(tmpDir, file.Filename)
	if err := c.SaveUploadedFile(file, tmpPath); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "cannot save uploaded file"})
		return
	}

	outputDir := filepath.Join(cfg.ProjectRoot, "public")
	result, err := converter.ConvertMaster(tmpPath, outputDir, cfg.PythonScript)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	outputs := make(map[string]string)
	if result.StudentsOutput != "" {
		outputs["students"] = result.StudentsOutput
	}
	if result.AttendanceOutput != "" {
		outputs["attendance"] = result.AttendanceOutput
	}
	if result.VedomostOutput != "" {
		outputs["vedomost"] = result.VedomostOutput
	}

	response := gin.H{
		"ok":      true,
		"message": "Мастер-конвертация завершена",
		"outputs": outputs,
	}

	if len(result.Warnings) > 0 {
		response["warnings"] = result.Warnings
	}
	if len(result.Errors) > 0 {
		response["errors"] = result.Errors
	}

	h.refreshHistory.AddEvent("success", "Мастер-конвертация выполнена через HTTP")

	c.JSON(http.StatusOK, response)
}

