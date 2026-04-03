package api

import (
	"net/http"

	"github.com/gin-gonic/gin"

	"dashboard/internal/services"
)

// DashboardMainHandler обрабатывает запросы для главного экрана (Real-time Dashboard)
type DashboardMainHandler struct {
	service *services.DashboardService
}

func NewDashboardMainHandler(service *services.DashboardService) *DashboardMainHandler {
	return &DashboardMainHandler{service: service}
}

// Stats возвращает общую статистику для главного экрана
// GET /api/dashboard/stats
func (h *DashboardMainHandler) Stats(c *gin.Context) {
	stats, err := h.service.GetStats()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	c.JSON(http.StatusOK, stats)
}

// TodayLessons возвращает список всех пар за сегодня
// GET /api/dashboard/lessons/today
func (h *DashboardMainHandler) TodayLessons(c *gin.Context) {
	lessons, err := h.service.GetTodayLessons()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	c.JSON(http.StatusOK, gin.H{
		"lessons": lessons,
		"count":   len(lessons),
	})
}

// CurrentLesson возвращает текущую пару (если есть)
// GET /api/dashboard/current-lesson
func (h *DashboardMainHandler) CurrentLesson(c *gin.Context) {
	current, err := h.service.GetCurrentLesson()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	c.JSON(http.StatusOK, current)
}
