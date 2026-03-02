package api

import (
	"net/http"
	"os"
	"path/filepath"

	"github.com/gin-gonic/gin"
)

// ServeSwagger обрабатывает запросы к Swagger UI и JSON
func ServeSwagger(c *gin.Context) {
	path := c.Param("path")

	if path == "/doc.json" {
		c.Header("Content-Type", "application/json")
		possiblePaths := []string{
			"docs/openapi.json",
			filepath.Join("backend", "docs", "openapi.json"),
		}
		var data []byte
		var err error
		for _, p := range possiblePaths {
			data, err = os.ReadFile(p)
			if err == nil {
				break
			}
		}
		if err != nil {
			data = []byte(`{"openapi":"3.0.0","info":{"title":"Dashboard API","version":"1.0"},"paths":{}}`)
		}
		c.Data(http.StatusOK, "application/json", data)
		return
	}

	if path == "" || path == "/" {
		html := `<!DOCTYPE html>
<html>
<head>
	<title>Dashboard API - Swagger UI</title>
	<link rel="stylesheet" type="text/css" href="https://unpkg.com/swagger-ui-dist@5.9.0/swagger-ui.css" />
	<style>
		html { box-sizing: border-box; overflow: -moz-scrollbars-vertical; overflow-y: scroll; }
		*, *:before, *:after { box-sizing: inherit; }
		body { margin:0; background: #fafafa; }
	</style>
</head>
<body>
	<div id="swagger-ui"></div>
	<script src="https://unpkg.com/swagger-ui-dist@5.9.0/swagger-ui-bundle.js"></script>
	<script src="https://unpkg.com/swagger-ui-dist@5.9.0/swagger-ui-standalone-preset.js"></script>
	<script>
		window.onload = function() {
			SwaggerUIBundle({
				url: "/swagger/doc.json",
				dom_id: '#swagger-ui',
				presets: [
					SwaggerUIBundle.presets.apis,
					SwaggerUIStandalonePreset
				],
				layout: "StandaloneLayout"
			});
		};
	</script>
</body>
</html>`
		c.Data(http.StatusOK, "text/html; charset=utf-8", []byte(html))
	} else {
		c.Status(http.StatusNotFound)
	}
}
