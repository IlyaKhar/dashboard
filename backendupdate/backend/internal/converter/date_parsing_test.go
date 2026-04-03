package converter

import "testing"

func TestParseDateValue_RussianMonth(t *testing.T) {
	t.Parallel()

	t.Run("2 апреля 2026", func(t *testing.T) {
		got := parseDateValue("2 апреля 2026")
		want := "2026-04-02"
		if got != want {
			t.Fatalf("parseDateValue: got=%q want=%q", got, want)
		}
	})

	t.Run("02.апреля.2026", func(t *testing.T) {
		got := parseDateValue("02.апреля.2026")
		want := "2026-04-02"
		if got != want {
			t.Fatalf("parseDateValue: got=%q want=%q", got, want)
		}
	})

	t.Run("3 апр. 26", func(t *testing.T) {
		got := parseDateValue("3 апр. 26")
		want := "2026-04-03"
		if got != want {
			t.Fatalf("parseDateValue: got=%q want=%q", got, want)
		}
	})
}

func TestParseDayAndOptionalDate_RussianMonth(t *testing.T) {
	t.Parallel()

	t.Run("Вторник 2 апреля 2026", func(t *testing.T) {
		isDay, dmy, iso := parseDayAndOptionalDate("Вторник 2 апреля 2026")
		if !isDay {
			t.Fatalf("parseDayAndOptionalDate: expected isDay=true")
		}
		if dmy != "02.04.2026" {
			t.Fatalf("parseDayAndOptionalDate: got dmy=%q want=%q", dmy, "02.04.2026")
		}
		if iso != "2026-04-02" {
			t.Fatalf("parseDayAndOptionalDate: got iso=%q want=%q", iso, "2026-04-02")
		}
	})

	t.Run("Вторник без даты", func(t *testing.T) {
		isDay, dmy, iso := parseDayAndOptionalDate("Вторник")
		if !isDay {
			t.Fatalf("parseDayAndOptionalDate: expected isDay=true")
		}
		if dmy != "" || iso != "" {
			t.Fatalf("parseDayAndOptionalDate: expected empty date, got dmy=%q iso=%q", dmy, iso)
		}
	})
}

func TestExtractLessonNumberFromRow_VedomostFormat(t *testing.T) {
	t.Parallel()

	got := extractLessonNumberFromRow([]string{"3", "", "", "", "", "10", "", "10"})
	if got != 3 {
		t.Fatalf("extractLessonNumberFromRow: got=%d want=%d", got, 3)
	}
}

func TestParseLessonNumberFromRow_VedomostFormat(t *testing.T) {
	t.Parallel()

	got := parseLessonNumberFromRow([]string{"4", "", "", "", "", "12", "", "12"})
	if got != 4 {
		t.Fatalf("parseLessonNumberFromRow: got=%d want=%d", got, 4)
	}
}

