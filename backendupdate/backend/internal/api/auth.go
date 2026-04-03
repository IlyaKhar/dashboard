package api

import (
	"net/http"
	"os"

	"github.com/gin-gonic/gin"
	"golang.org/x/crypto/bcrypt"

	"dashboard/internal/config"
	"dashboard/internal/middleware"
)

// AuthHandler обрабатывает запросы авторизации
type AuthHandler struct {
	cfg *config.Config
}

// NewAuthHandler создаёт новый handler авторизации
func NewAuthHandler(cfg *config.Config) *AuthHandler {
	return &AuthHandler{cfg: cfg}
}

// Login обрабатывает POST /api/login
func (h *AuthHandler) Login(c *gin.Context) {
	var body struct {
		Username string `json:"username"`
		Password string `json:"password"`
	}

	if err := c.ShouldBindJSON(&body); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid JSON"})
		return
	}

	// Проверяем учётные данные.
	if body.Username != h.cfg.LoginUser {
		c.JSON(http.StatusUnauthorized, gin.H{"error": "Unauthorized"})
		return
	}

	// В production LOGIN_PASSWORD хранится как bcrypt-хэш.
	// В dev-окружении допускаем fallback на простой текст для удобства.
	if err := bcrypt.CompareHashAndPassword([]byte(h.cfg.LoginPassword), []byte(body.Password)); err != nil {
		if os.Getenv("APP_ENV") != "production" && h.cfg.LoginPassword == body.Password {
			// dev-режим, поддержка старого plain-text пароля
		} else {
			c.JSON(http.StatusUnauthorized, gin.H{"error": "Unauthorized"})
			return
		}
	}

	// Генерируем JWT токен
	token, err := middleware.IssueJWT(h.cfg.JWTSecret, h.cfg.LoginRole)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Internal error"})
		return
	}

	c.JSON(http.StatusOK, gin.H{
		"token": token,
		"role":  h.cfg.LoginRole,
	})
}
