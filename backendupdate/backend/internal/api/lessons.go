package api

import (
	"net/http"
	"time"

	"github.com/gin-gonic/gin"

	"dashboard/internal/services"
)

// LessonsHandler обрабатывает запросы по расписанию занятий
type LessonsHandler struct {
	service *services.LessonsService
}

func NewLessonsHandler(service *services.LessonsService) *LessonsHandler {
	return &LessonsHandler{service: service}
}

// Day возвращает список занятий за указанный день
// GET /api/lessons/day?date=2026-02-02
// если date не указан — берём сегодня
func (h *LessonsHandler) Day(c *gin.Context) {
	dateStr := c.Query("date")

	var (
		targetDate time.Time
		err        error
	)

	if dateStr == "" {
		targetDate = time.Now()
	} else {
		targetDate, err = time.Parse("2006-01-02", dateStr)
		if err != nil {
			c.JSON(http.StatusBadRequest, gin.H{"error": "invalid date format, expected YYYY-MM-DD"})
			return
		}
	}

	lessons, err := h.service.GetDayLessons(targetDate)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	c.JSON(http.StatusOK, gin.H{
		"date":    targetDate.Format("2006-01-02"),
		"lessons": lessons,
	})
}

