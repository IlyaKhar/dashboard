package api

import (
	"github.com/gin-gonic/gin"
	"dashboard/internal/config"
	"dashboard/internal/database"
	"dashboard/internal/middleware"
	"dashboard/internal/scheduler"
	"dashboard/internal/services"
)

// SetupRouter настраивает и возвращает Gin router
func SetupRouter(
	cfg *config.Config,
	sched *scheduler.Scheduler,
	dbLoader *database.Loader,
	attendanceService *services.AttendanceService,
	scheduleService *services.ScheduleService,
	reconciliationService *services.ReconciliationService,
	lessonsService *services.LessonsService,
	dashboardMainService *services.DashboardService,
	thresholdsService *services.ThresholdsService,
	loginRateLimiter *middleware.RateLimiter,
	refreshHistory *RefreshHistoryStore,
) *gin.Engine {
	router := gin.New()

	router.Use(middleware.SetupCORS())
	router.Use(middleware.Logger())
	router.Use(middleware.Recovery())

	router.Use(func(c *gin.Context) {
		c.Set("attendance_output", cfg.AttendanceOutput)
		c.Set("statement_output", cfg.StatementOutput)
		c.Next()
	})

	ginHandler := NewGinHandler(sched, dbLoader, reconciliationService, refreshHistory)
	authHandler := NewAuthHandler(cfg)
	dashboardHandler := NewDashboardHandler(attendanceService, cfg.AbsenceThreshold)
	reconcileHandler := NewReconciliationHandler(reconciliationService)
	lessonsHandler := NewLessonsHandler(lessonsService)
	dashboardMainHandler := NewDashboardMainHandler(dashboardMainService)
	thresholdsHandler := NewThresholdsHandler(thresholdsService)

	apiGroup := router.Group("/api")
	{
		apiGroup.POST("/login", middleware.LoginRateLimit(loginRateLimiter), authHandler.Login)
		apiGroup.GET("/health", ginHandler.HealthCheck)

		protected := apiGroup.Group("")
		protected.Use(middleware.JWTAuth(cfg.JWTSecret))
		{
			protected.GET("/attendance", dashboardHandler.List)
			protected.GET("/attendance/summary", dashboardHandler.Summary)
			protected.GET("/attendance/drill/departments", dashboardHandler.DrillDepartments)
			protected.GET("/attendance/drill/groups", dashboardHandler.DrillGroups)
			protected.GET("/attendance/drill/students", dashboardHandler.DrillStudents)
			protected.GET("/attendance/reconcile/day", reconcileHandler.ReconcileDay)
			protected.GET("/attendance/reconcile/day/group", reconcileHandler.ReconcileDayLessonGroup)
			protected.GET("/lessons/day", lessonsHandler.Day)
			protected.GET("/dashboard/stats", dashboardMainHandler.Stats)
			protected.GET("/dashboard/lessons/today", dashboardMainHandler.TodayLessons)
			protected.GET("/dashboard/current-lesson", dashboardMainHandler.CurrentLesson)
			protected.GET("/settings/thresholds", thresholdsHandler.GetThresholds)

			adminGroup := protected.Group("/admin")
			adminGroup.Use(middleware.RequireRole("admin"))
			{
				adminGroup.POST("/refresh-data", ginHandler.RefreshData)
				adminGroup.GET("/refresh-status", ginHandler.GetRefreshStatus)
				adminGroup.GET("/refresh-history", ginHandler.GetRefreshHistory)
				adminGroup.POST("/convert/statement", ginHandler.ConvertStatement)
				adminGroup.POST("/convert/schedule", ginHandler.ConvertSchedule)
				adminGroup.POST("/convert/master", ginHandler.ConvertMaster)
			}

			settingsGroup := protected.Group("/settings")
			settingsGroup.Use(middleware.RequireRole("admin"))
			{
				settingsGroup.PUT("/thresholds", thresholdsHandler.UpdateThresholds)
			}
		}
	}

	router.GET("/swagger/*path", ServeSwagger)

	return router
}
