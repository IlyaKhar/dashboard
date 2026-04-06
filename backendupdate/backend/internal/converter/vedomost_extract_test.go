package converter

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// Проверка на реальном шаблоне: строка «3»/«3.0» не должна становиться группой; номер пары должен доходить до строк студентов.
func TestExtractAllData_VedomostLessonRowNotGroup(t *testing.T) {
	dir := t.TempDir()
	xls := filepath.Join("ведомость.xls")
	if _, err := os.Stat(xls); err != nil {
		t.Skip("нет локального ведомость.xls в каталоге converter")
	}
	xlsx := filepath.Join(dir, "v.xlsx")
	if err := convertXLSToXLSX(xls, xlsx, filepath.Join(filepath.Dir(xls), "xls_to_xlsx.py")); err != nil {
		t.Fatalf("xls→xlsx: %v", err)
	}
	f, err := excelize.OpenFile(xlsx)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	sn := f.GetSheetName(0)
	rows, err := f.GetRows(sn)
	if err != nil {
		t.Fatal(err)
	}
	data := extractAllData(rows)

	var badGroups []string
	for _, r := range data.AttendanceRecords {
		if r.Group == "3.0" || r.Group == "2.0" || r.Group == "1.0" {
			badGroups = append(badGroups, r.Group)
		}
	}
	if len(badGroups) > 0 {
		t.Fatalf("в attendance попали «группы»-числа: %v", badGroups)
	}

	// После фикса isGroup хотя бы одна запись с парой ≥2 должна быть (в файле есть блоки с «4.0» и т.д.)
	maxLn := 0
	for _, r := range data.AttendanceRecords {
		if r.LessonNumber > maxLn {
			maxLn = r.LessonNumber
		}
	}
	if maxLn < 2 {
		t.Fatalf("ожидали номера пар до 2+, получили max lessonNumber=%d (возможно сломана строка номера пары)", maxLn)
	}
}
