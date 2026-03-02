package api

import (
	"net/http"

	"github.com/gin-gonic/gin"

	"dashboard/internal/services"
)

// ThresholdsHandler обрабатывает запросы для управления порогами
type ThresholdsHandler struct {
	service *services.ThresholdsService
}

func NewThresholdsHandler(service *services.ThresholdsService) *ThresholdsHandler {
	return &ThresholdsHandler{service: service}
}

// GetThresholds возвращает текущие пороги
// GET /api/settings/thresholds?type=lesson
func (h *ThresholdsHandler) GetThresholds(c *gin.Context) {
	thresholdType := c.DefaultQuery("type", "lesson")

	thresholds, err := h.service.GetThresholds(thresholdType)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	c.JSON(http.StatusOK, thresholds)
}

// UpdateThresholds обновляет пороги (только для администраторов)
// PUT /api/settings/thresholds
func (h *ThresholdsHandler) UpdateThresholds(c *gin.Context) {
	var req struct {
		Type           string  `json:"type" binding:"required"`
		GreenThreshold float64 `json:"green_threshold" binding:"required"`
		YellowThreshold float64 `json:"yellow_threshold" binding:"required"`
	}

	if err := c.ShouldBindJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	if err := h.service.UpdateThresholds(req.Type, req.GreenThreshold, req.YellowThreshold); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	c.JSON(http.StatusOK, gin.H{"message": "Пороги успешно обновлены"})
}
