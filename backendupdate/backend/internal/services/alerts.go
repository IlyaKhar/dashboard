package services

import (
	"log"

	"dashboard/internal/models"
)

// CheckAlerts проверяет пороги пропусков и логирует предупреждения
func CheckAlerts(data []models.FlatRecord, threshold int) {
	groupMissed := make(map[string]int)
	groupStudents := make(map[string]map[string]struct{})

	// Собираем статистику по группам
	for _, rec := range data {
		groupMissed[rec.Group] += rec.Missed
		if groupStudents[rec.Group] == nil {
			groupStudents[rec.Group] = make(map[string]struct{})
		}
		groupStudents[rec.Group][rec.Student] = struct{}{}
	}

	// Проверяем пороги
	for grp, missed := range groupMissed {
		n := len(groupStudents[grp])
		if n == 0 {
			continue
		}
		avg := missed / n
		if avg >= threshold {
			log.Printf("[Alerts] ALERT: группа %s превысила порог пропусков: среднее %d часов на студента (порог: %d)", grp, avg, threshold)
		}
	}
}
