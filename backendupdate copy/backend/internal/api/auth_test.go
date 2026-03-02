package api

import (
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/gin-gonic/gin"
	"dashboard/internal/config"
)

func TestAuthHandler_Login_InvalidJSON(t *testing.T) {
	gin.SetMode(gin.TestMode)

	cfg := &config.Config{
		LoginUser:     "admin",
		LoginPassword: "secret",
		LoginRole:     "admin",
		JWTSecret:     "test-secret",
	}
	h := NewAuthHandler(cfg)
	r := gin.New()
	r.POST("/login", h.Login)

	req := httptest.NewRequest(http.MethodPost, "/login", strings.NewReader("invalid json"))
	req.Header.Set("Content-Type", "application/json")
	w := httptest.NewRecorder()
	r.ServeHTTP(w, req)

	if w.Code != http.StatusBadRequest {
		t.Errorf("ожидался 400, получен %d", w.Code)
	}
}

func TestAuthHandler_Login_Unauthorized(t *testing.T) {
	gin.SetMode(gin.TestMode)

	cfg := &config.Config{
		LoginUser:     "admin",
		LoginPassword: "secret",
		LoginRole:     "admin",
		JWTSecret:     "test-secret",
	}
	h := NewAuthHandler(cfg)
	r := gin.New()
	r.POST("/login", h.Login)

	body := `{"username":"admin","password":"wrong"}`
	req := httptest.NewRequest(http.MethodPost, "/login", strings.NewReader(body))
	req.Header.Set("Content-Type", "application/json")
	w := httptest.NewRecorder()
	r.ServeHTTP(w, req)

	if w.Code != http.StatusUnauthorized {
		t.Errorf("ожидался 401, получен %d", w.Code)
	}
}

func TestAuthHandler_Login_Success(t *testing.T) {
	gin.SetMode(gin.TestMode)

	cfg := &config.Config{
		LoginUser:     "admin",
		LoginPassword: "secret",
		LoginRole:     "admin",
		JWTSecret:     "test-secret",
	}
	h := NewAuthHandler(cfg)
	r := gin.New()
	r.POST("/login", h.Login)

	body := `{"username":"admin","password":"secret"}`
	req := httptest.NewRequest(http.MethodPost, "/login", strings.NewReader(body))
	req.Header.Set("Content-Type", "application/json")
	w := httptest.NewRecorder()
	r.ServeHTTP(w, req)

	if w.Code != http.StatusOK {
		t.Errorf("ожидался 200, получен %d: %s", w.Code, w.Body.String())
	}
	if !strings.Contains(w.Body.String(), "token") {
		t.Errorf("ожидался token в ответе: %s", w.Body.String())
	}
}
