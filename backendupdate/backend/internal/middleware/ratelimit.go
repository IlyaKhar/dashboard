package middleware

import (
	"net/http"
	"sync"
	"time"

	"github.com/gin-gonic/gin"
)

// RateLimiter простой in-memory rate limiter для защиты /login
type RateLimiter struct {
	mu       sync.Mutex
	attempts map[string][]time.Time
	window   time.Duration
	maxReq   int
}

// NewRateLimiter создаёт лимитер: maxReq запросов за window с одного IP
func NewRateLimiter(maxReq int, window time.Duration) *RateLimiter {
	return &RateLimiter{
		attempts: make(map[string][]time.Time),
		window:   window,
		maxReq:   maxReq,
	}
}

func (r *RateLimiter) cleanup(key string) {
	cutoff := time.Now().Add(-r.window)
	attempts := r.attempts[key]
	var valid []time.Time
	for _, t := range attempts {
		if t.After(cutoff) {
			valid = append(valid, t)
		}
	}
	if len(valid) == 0 {
		delete(r.attempts, key)
	} else {
		r.attempts[key] = valid
	}
}

func (r *RateLimiter) Allow(key string) bool {
	r.mu.Lock()
	defer r.mu.Unlock()

	r.cleanup(key)
	attempts := r.attempts[key]
	if len(attempts) >= r.maxReq {
		return false
	}
	r.attempts[key] = append(attempts, time.Now())
	return true
}

// LoginRateLimit middleware — 5 попыток в минуту с одного IP
func LoginRateLimit(limiter *RateLimiter) gin.HandlerFunc {
	return func(c *gin.Context) {
		ip := c.ClientIP()
		if !limiter.Allow(ip) {
			c.JSON(http.StatusTooManyRequests, gin.H{
				"error": "Слишком много попыток входа. Попробуйте через минуту.",
			})
			c.Abort()
			return
		}
		c.Next()
	}
}
