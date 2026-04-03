package api

import (
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/gin-gonic/gin"

	"dashboard/internal/services"
)

// ReconciliationHandler обрабатывает запросы сверки расписания и посещаемости.
type ReconciliationHandler struct {
	service *services.ReconciliationService
}

func NewReconciliationHandler(service *services.ReconciliationService) *ReconciliationHandler {
	return &ReconciliationHandler{service: service}
}

// ReconcileDay выполняет сверку за конкретную дату.
// GET /api/attendance/reconcile/day?date=YYYY-MM-DD
// С опциональным параметром lesson=1..6 — сверка только по этой паре (для сценария «первая пара — кто где»).
func (h *ReconciliationHandler) ReconcileDay(c *gin.Context) {
	date := strings.TrimSpace(c.Query("date"))
	if date == "" {
		date = time.Now().Format("2006-01-02")
	}
	if _, err := time.Parse("2006-01-02", date); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "invalid date, expected YYYY-MM-DD"})
		return
	}

	lessonStr := strings.TrimSpace(c.Query("lesson"))
	lesson := 0
	if lessonStr != "" {
		var err error
		lesson, err = strconv.Atoi(lessonStr)
		if err != nil || lesson < 1 || lesson > 6 {
			c.JSON(http.StatusBadRequest, gin.H{"error": "lesson must be 1-6"})
			return
		}
	}

	var result *services.ReconcileResult
	var err error
	if lesson > 0 {
		result, err = h.service.ReconcileDayLesson(date, lesson)
	} else {
		result, err = h.service.ReconcileDay(date)
	}
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	c.JSON(http.StatusOK, result)
}

// ReconcileDayLessonGroup — по группе на пару: дисциплина, кто есть, кто нет.
// GET /api/attendance/reconcile/day/group?date=YYYY-MM-DD&lesson=1&group=1вб1
func (h *ReconciliationHandler) ReconcileDayLessonGroup(c *gin.Context) {
	date := strings.TrimSpace(c.Query("date"))
	if date == "" {
		date = time.Now().Format("2006-01-02")
	}
	if _, err := time.Parse("2006-01-02", date); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "invalid date, expected YYYY-MM-DD"})
		return
	}
	lessonStr := strings.TrimSpace(c.Query("lesson"))
	lesson := 1
	if lessonStr != "" {
		var err error
		lesson, err = strconv.Atoi(lessonStr)
		if err != nil || lesson < 1 || lesson > 6 {
			c.JSON(http.StatusBadRequest, gin.H{"error": "lesson must be 1-6"})
			return
		}
	}
	group := strings.TrimSpace(c.Query("group"))
	if group == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "group is required"})
		return
	}

	detail, err := h.service.ReconcileDayLessonGroup(date, lesson, group)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	if detail == nil {
		c.JSON(http.StatusOK, gin.H{
			"group": group, "message": "на эту пару у группы нет занятия",
		})
		return
	}
	c.JSON(http.StatusOK, detail)
}

