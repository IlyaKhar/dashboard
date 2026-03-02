package api

import (
	"encoding/json"
	"log"
	"net/http"
	"time"

	"dashboard/internal/database"
	"dashboard/internal/scheduler"
)

// @title Dashboard Backend API
// @version 1.0
// @description API для управления дашбордом посещаемости студентов
// @termsOfService http://swagger.io/terms/

// @contact.name API Support
// @contact.email support@dashboard.local

// @host localhost:8080
// @BasePath /api

// Handler содержит обработчики API
type Handler struct {
	scheduler         *scheduler.Scheduler
	dbLoader          *database.Loader
	lastRefresh       time.Time
	refreshInProgress bool
}

func NewHandler(scheduler *scheduler.Scheduler, dbLoader *database.Loader) *Handler {
	return &Handler{
		scheduler: scheduler,
		dbLoader:  dbLoader,
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
func (h *Handler) RefreshData(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Метод не разрешён", http.StatusMethodNotAllowed)
		return
	}

	// Проверяем, не выполняется ли уже обновление
	if h.refreshInProgress {
		respondJSON(w, http.StatusConflict, map[string]string{
			"error": "Обновление уже выполняется",
		})
		return
	}

	h.refreshInProgress = true
	defer func() {
		h.refreshInProgress = false
	}()

	log.Println("[API] Запуск ручного обновления данных...")

	// Запускаем обновление
	if err := h.scheduler.RefreshData(); err != nil {
		log.Printf("[API] Ошибка обновления данных: %v", err)
		respondJSON(w, http.StatusInternalServerError, map[string]string{
			"error":   "Ошибка обновления данных",
			"details": err.Error(),
		})
		return
	}

	// Загружаем в БД (если БД подключена)
	// Пути будут передаваться через конфиг позже, пока используем относительные
	attendancePath := "../../public/attendance.json"
	statementPath := "../../public/vedomost.json"
	
	if err := h.dbLoader.LoadAttendance(attendancePath); err != nil {
		log.Printf("[API] Предупреждение при загрузке посещаемости в БД: %v", err)
	}
	if err := h.dbLoader.LoadStatement(statementPath); err != nil {
		log.Printf("[API] Предупреждение при загрузке ведомости в БД: %v", err)
	}

	h.lastRefresh = time.Now()

	respondJSON(w, http.StatusOK, map[string]interface{}{
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
func (h *Handler) GetRefreshStatus(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "Метод не разрешён", http.StatusMethodNotAllowed)
		return
	}

	status := map[string]interface{}{
		"in_progress": h.refreshInProgress,
	}

	if !h.lastRefresh.IsZero() {
		status["last_refresh"] = h.lastRefresh.Format(time.RFC3339)
		status["last_refresh_ago"] = time.Since(h.lastRefresh).String()
	} else {
		status["last_refresh"] = nil
		status["last_refresh_ago"] = nil
	}

	respondJSON(w, http.StatusOK, status)
}

// GetRefreshHistory возвращает историю обновлений (пока заглушка)
// @Summary История обновлений
// @Description Возвращает историю обновлений данных
// @Tags admin
// @Produce json
// @Success 200 {object} map[string]interface{} "История обновлений"
// @Router /admin/refresh-history [get]
func (h *Handler) GetRefreshHistory(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "Метод не разрешён", http.StatusMethodNotAllowed)
		return
	}

	// TODO: Реальная история из БД
	history := []map[string]interface{}{
		{
			"time":    h.lastRefresh.Format(time.RFC3339),
			"status":  "success",
			"message": "Обновление выполнено успешно",
		},
	}

	respondJSON(w, http.StatusOK, map[string]interface{}{
		"history": history,
	})
}

// HealthCheck проверяет работоспособность сервера
// @Summary Health Check
// @Description Проверяет работоспособность сервера
// @Tags system
// @Produce json
// @Success 200 {object} map[string]string "Сервер работает"
// @Router /health [get]
func (h *Handler) HealthCheck(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "Метод не разрешён", http.StatusMethodNotAllowed)
		return
	}

	respondJSON(w, http.StatusOK, map[string]string{
		"status":  "ok",
		"service": "dashboard-backend",
	})
}

// respondJSON отправляет JSON ответ
func respondJSON(w http.ResponseWriter, statusCode int, data interface{}) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(statusCode)
	if err := json.NewEncoder(w).Encode(data); err != nil {
		log.Printf("[API] Ошибка кодирования JSON: %v", err)
	}
}
