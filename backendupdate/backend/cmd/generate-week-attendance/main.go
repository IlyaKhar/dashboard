// Генератор тестовых данных посещаемости на 7 дней (учебная неделя).
// Контингент из students.json; пропуски — либо из эталонного attendance (-ref),
// либо по реалистичному распределению. Даты недели задаются -week.
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"math/rand"
	"os"
	"path/filepath"
	"time"
)

type attendanceRecord struct {
	Date   string `json:"date"`
	Missed int    `json:"missed"`
}

type studentAttendance struct {
	Student    string             `json:"student"`
	Attendance []attendanceRecord `json:"attendance"`
}

type groupAttendance struct {
	Group    string              `json:"group"`
	Students []studentAttendance  `json:"students"`
}

type departmentAttendance struct {
	Department string            `json:"department"`
	Groups     []groupAttendance `json:"groups"`
}

type studentsRoot struct {
	Departments []struct {
		Department string `json:"department"`
		Groups     []struct {
			Group    string `json:"group"`
			Students []struct {
				FullName string `json:"fullName"`
			} `json:"students"`
		} `json:"groups"`
	} `json:"departments"`
}

// distribution хранит вероятности для missed: 0, 2, 4, 6, 8 (индексы 0,1,2,3,4)
type distribution struct {
	weights [5]int // 0, 2, 4, 6, 8
	total   int
}

func (d *distribution) pick(r *rand.Rand) int {
	vals := [5]int{0, 2, 4, 6, 8}
	if d.total == 0 {
		return 0
	}
	n := r.Intn(d.total)
	for i, w := range d.weights {
		n -= w
		if n < 0 {
			return vals[i]
		}
	}
	return 0
}

// buildDistribution из эталонного attendance собирает распределение пропусков (0, 2, 4, 6, 8)
func buildDistribution(ref []departmentAttendance) distribution {
	var d distribution
	vals := map[int]int{0: 0, 2: 1, 4: 2, 6: 3, 8: 4}
	for _, dept := range ref {
		for _, gr := range dept.Groups {
			for _, st := range gr.Students {
				for _, rec := range st.Attendance {
					v := rec.Missed
					if v > 8 {
						v = 8
					}
					if idx, ok := vals[v]; ok {
						d.weights[idx]++
						d.total++
					} else {
						d.weights[0]++
						d.total++
					}
				}
			}
		}
	}
	// Если в эталоне почти нет 0 (только те, у кого были пропуски), добавляем долю нулей
	if d.total > 0 && d.weights[0] < d.total/2 {
		// ~80% дней без пропусков в реалистичной неделе
		d.weights[0] = d.total * 4
		d.total += d.total * 4
	}
	return d
}

func main() {
	rng := rand.New(rand.NewSource(time.Now().UnixNano()))

	outDir := flag.String("out", "", "Директория public (attendance.json и students.json)")
	weekStart := flag.String("week", "2026-02-16", "Понедельник недели (YYYY-MM-DD)")
	refPath := flag.String("ref", "", "Путь к эталонному attendance.json (15.01, 18.02 и т.д.) — по нему берётся распределение пропусков")
	flag.Parse()

	if *outDir == "" {
		*outDir = filepath.Join("..", "..", "..", "public")
	}
	attPath := filepath.Join(*outDir, "attendance.json")
	stuPath := filepath.Join(*outDir, "students.json")

	// Распределение пропусков: из эталона или дефолтное
	var dist distribution
	if *refPath != "" {
		refData, err := os.ReadFile(*refPath)
		if err != nil {
			log.Fatalf("Чтение эталона %s: %v", *refPath, err)
		}
		var ref []departmentAttendance
		if err := json.Unmarshal(refData, &ref); err != nil {
			log.Fatalf("Парсинг эталона: %v", err)
		}
		dist = buildDistribution(ref)
		log.Printf("Эталон: использовано распределение пропусков (всего записей %d)", dist.total)
	}
	if dist.total == 0 {
		// Дефолт: не «всё зелёное» — есть доля с пропусками
		dist = distribution{
			weights: [5]int{820, 100, 50, 20, 10},
			total:   1000,
		}
	}

	// Загружаем students.json
	studentsData, err := os.ReadFile(stuPath)
	if err != nil {
		log.Fatalf("Чтение students.json: %v", err)
	}
	var students studentsRoot
	if err := json.Unmarshal(studentsData, &students); err != nil {
		log.Fatalf("Парсинг students.json: %v", err)
	}

	start, err := time.Parse("2006-01-02", *weekStart)
	if err != nil {
		log.Fatalf("Неверная дата недели %q: %v", *weekStart, err)
	}
	dates := make([]string, 7)
	for i := 0; i < 7; i++ {
		dates[i] = start.AddDate(0, 0, i).Format("2006-01-02")
	}

	var result []departmentAttendance
	for _, d := range students.Departments {
		dept := departmentAttendance{Department: d.Department, Groups: nil}
		for _, g := range d.Groups {
			gr := groupAttendance{Group: g.Group, Students: nil}
			for _, s := range g.Students {
				recs := make([]attendanceRecord, 7)
				for i, date := range dates {
					recs[i] = attendanceRecord{
						Date:   date,
						Missed: dist.pick(rng),
					}
				}
				gr.Students = append(gr.Students, studentAttendance{
					Student:    s.FullName,
					Attendance: recs,
				})
			}
			dept.Groups = append(dept.Groups, gr)
		}
		result = append(result, dept)
	}

	raw, err := json.MarshalIndent(result, "", "  ")
	if err != nil {
		log.Fatalf("Сборка JSON: %v", err)
	}
	if err := os.WriteFile(attPath, raw, 0644); err != nil {
		log.Fatalf("Запись %s: %v", attPath, err)
	}

	fmt.Printf("✅ Сгенерирована посещаемость на 7 дней (%s — %s)\n", dates[0], dates[6])
	fmt.Printf("   Файл: %s\n", attPath)
	if *refPath != "" {
		fmt.Printf("   Распределение пропусков: по эталону %s\n", *refPath)
	}
}
