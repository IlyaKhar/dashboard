package services

import (
	"database/sql"
	"fmt"

	"dashboard/internal/models"
)

// ThresholdsService предоставляет бизнес-логику для работы с порогами
type ThresholdsService struct {
	db *sql.DB
}

func NewThresholdsService(db *sql.DB) *ThresholdsService {
	return &ThresholdsService{db: db}
}

// GetThresholds возвращает пороги для указанного типа
func (s *ThresholdsService) GetThresholds(thresholdType string) (*models.Threshold, error) {
	if s.db == nil {
		return nil, fmt.Errorf("БД не подключена")
	}

	var t models.Threshold
	err := s.db.QueryRow(`
		SELECT id, type, green_threshold, yellow_threshold, created_at, updated_at
		FROM thresholds
		WHERE type = $1
	`, thresholdType).Scan(
		&t.ID,
		&t.Type,
		&t.GreenThreshold,
		&t.YellowThreshold,
		&t.CreatedAt,
		&t.UpdatedAt,
	)

	if err == sql.ErrNoRows {
		// Если порогов нет - возвращаем дефолтные
		return &models.Threshold{
			Type:           thresholdType,
			GreenThreshold: 90.0,
			YellowThreshold: 70.0,
		}, nil
	}
	if err != nil {
		return nil, fmt.Errorf("ошибка получения порогов: %w", err)
	}

	return &t, nil
}

// UpdateThresholds обновляет пороги для указанного типа
func (s *ThresholdsService) UpdateThresholds(thresholdType string, greenThreshold, yellowThreshold float64) error {
	if s.db == nil {
		return fmt.Errorf("БД не подключена")
	}

	// Проверяем валидность порогов
	if greenThreshold < yellowThreshold {
		return fmt.Errorf("зелёный порог должен быть >= жёлтого")
	}
	if greenThreshold < 0 || greenThreshold > 100 {
		return fmt.Errorf("зелёный порог должен быть от 0 до 100")
	}
	if yellowThreshold < 0 || yellowThreshold > 100 {
		return fmt.Errorf("жёлтый порог должен быть от 0 до 100")
	}

	_, err := s.db.Exec(`
		INSERT INTO thresholds (type, green_threshold, yellow_threshold)
		VALUES ($1, $2, $3)
		ON CONFLICT (type) 
		DO UPDATE SET 
			green_threshold = EXCLUDED.green_threshold,
			yellow_threshold = EXCLUDED.yellow_threshold,
			updated_at = CURRENT_TIMESTAMP
	`, thresholdType, greenThreshold, yellowThreshold)

	if err != nil {
		return fmt.Errorf("ошибка обновления порогов: %w", err)
	}

	return nil
}
