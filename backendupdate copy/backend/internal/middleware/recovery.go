package middleware

import (
	"log"
	"net/http"

	"github.com/gin-gonic/gin"
)

// Recovery middleware для обработки паник
func Recovery() gin.HandlerFunc {
	return gin.CustomRecovery(func(c *gin.Context, recovered interface{}) {
		log.Printf("[Recovery] Паника: %v", recovered)
		c.JSON(http.StatusInternalServerError, gin.H{
			"error":   "Внутренняя ошибка сервера",
			"message": "Произошла непредвиденная ошибка",
		})
		c.Abort()
	})
}
