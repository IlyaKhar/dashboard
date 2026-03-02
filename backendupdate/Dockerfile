### Стейдж сборки Go-бэкенда
FROM golang:1.24-alpine AS builder

WORKDIR /app/backend

RUN apk add --no-cache git

# Модульные файлы
COPY backend/go.mod backend/go.sum ./
RUN go mod download

# Весь исходник бэкенда
COPY backend/. .

# Собираем бинарник сервера
RUN CGO_ENABLED=0 GOOS=linux go build -o server ./cmd/server


### Рантайм-стейдж
FROM alpine:3.20

WORKDIR /app

RUN apk add --no-cache ca-certificates && update-ca-certificates

# Бинарник сервера
COPY --from=builder /app/backend/server ./backend/server

# Папка с JSON и статикой (attendance.json, schedule.json, oksei_timetable.json и т.д.)
COPY public ./public

# Настройки сервера по умолчанию
ENV SERVER_HOST=0.0.0.0 \
    SERVER_PORT=8080

EXPOSE 8080

# Запускаем из папки backend, чтобы config.Load увидел корень /app (там есть backend/ и public/)
WORKDIR /app/backend

CMD ["./server"]

