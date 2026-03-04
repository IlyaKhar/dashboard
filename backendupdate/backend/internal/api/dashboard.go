package api

import (
	"net/http"
	"strings"

	"dashboard/internal/models"
	"dashboard/internal/pkg/validator"
	"dashboard/internal/services"

	"github.com/gin-gonic/gin"
)

type DashboardHandler struct {
	attendanceService *services.AttendanceService
	alertsThreshold   int
}

func NewDashboardHandler(attendanceService *services.AttendanceService, alertsThreshold int) *DashboardHandler {
	return &DashboardHandler{
		attendanceService: attendanceService,
		alertsThreshold:   alertsThreshold,
	}
}

func (h *DashboardHandler) List(c *gin.Context) {
	_, flat, err := h.attendanceService.LoadFromJSON()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot load attendance"})
		return
	}

	params := services.ParseFilterParams(c.Request)
	filtered := h.attendanceService.Filter(flat, params)

	services.CheckAlerts(filtered, h.alertsThreshold)

	c.JSON(http.StatusOK, filtered)
}

func (h *DashboardHandler) Summary(c *gin.Context) {
	params := services.ParseFilterParams(c.Request)

	// Если запрошен конкретный день / диапазон / период — используем накопительную историю
	// (attendance_history.json). Без фильтров — оперативный режим по последнему дню.
	var (
		departments []models.DepartmentJSON
		flat        []models.FlatRecord
		err         error
	)

	if params.Date != "" || params.DateFrom != "" || params.DateTo != "" || params.Period != "" {
		departments, flat, err = h.attendanceService.LoadHistoryJSON()
	} else {
		departments, flat, err = h.attendanceService.LoadFromJSON()
	}
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot load attendance"})
		return
	}

	filtered := h.attendanceService.Filter(flat, params)
	summary := h.attendanceService.BuildSummary(departments, filtered)

	c.JSON(http.StatusOK, summary)
}

func (h *DashboardHandler) DrillDepartments(c *gin.Context) {
	departments, flat, err := h.attendanceService.LoadFromJSON()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot load attendance"})
		return
	}

	params := services.ParseFilterParams(c.Request)
	filtered := h.attendanceService.Filter(flat, params)
	result := h.attendanceService.BuildDrillDepartments(departments, filtered)

	c.JSON(http.StatusOK, result)
}

func (h *DashboardHandler) DrillGroups(c *gin.Context) {
	department, ok := validator.ValidateDepartment(strings.TrimSpace(c.Query("department")))
	if !ok || department == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "department parameter required and must be valid"})
		return
	}

	departments, flat, err := h.attendanceService.LoadFromJSON()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot load attendance"})
		return
	}

	params := services.ParseFilterParams(c.Request)
	filtered := h.attendanceService.Filter(flat, params)
	result := h.attendanceService.BuildDrillGroups(departments, filtered, department)

	c.JSON(http.StatusOK, result)
}

func (h *DashboardHandler) DrillStudents(c *gin.Context) {
	department, okDept := validator.ValidateDepartment(strings.TrimSpace(c.Query("department")))
	group, okGrp := validator.ValidateGroup(strings.TrimSpace(c.Query("group")))

	if !okDept || !okGrp || department == "" || group == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "department and group parameters required and must be valid"})
		return
	}

	_, flat, err := h.attendanceService.LoadFromJSON()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot load attendance"})
		return
	}

	params := services.ParseFilterParams(c.Request)
	filtered := h.attendanceService.Filter(flat, params)
	result := h.attendanceService.BuildDrillStudents(filtered, department, group)

	c.JSON(http.StatusOK, result)
}
