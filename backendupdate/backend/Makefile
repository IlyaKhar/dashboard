.PHONY: swagger
swagger:
	@echo "Генерация Swagger документации..."
	@swag init -g cmd/server/main.go -o ./docs
	@echo "✅ Swagger документация сгенерирована в ./docs"

.PHONY: run
run:
	@go run ./cmd/server

.PHONY: build
build:
	@go build -o bin/server ./cmd/server

.PHONY: test
test:
	@go test ./... -v
