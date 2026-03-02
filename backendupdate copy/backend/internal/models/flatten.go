package models

// VedomostToFlat конвертирует vedomost.json в плоский список для сверки: один запись на студента с пропусками (date подставляется).
func VedomostToFlat(vedomost []VedomostDepartment, date string) []FlatRecord {
	var out []FlatRecord
	for _, d := range vedomost {
		for _, spec := range d.Specialties {
			for _, g := range spec.Groups {
				for _, st := range g.Students {
					if st.MissedTotal <= 0 {
						continue
					}
					out = append(out, FlatRecord{
						Department: d.Department,
						Group:      g.Group,
						Student:    st.Student,
						Date:       date,
						Missed:     st.MissedTotal,
					})
				}
			}
		}
	}
	return out
}

// FlatRecord представляет плоскую запись посещаемости
type FlatRecord struct {
	Department string `json:"department"`
	Group      string `json:"group"`
	Student    string `json:"student"`
	Date       string `json:"date"`
	Missed     int    `json:"missed"`
}

// Flatten преобразует иерархию DepartmentJSON → GroupJSON → StudentJSON → AttendanceRecordJSON
// в плоский список записей для удобной фильтрации и поиска
func Flatten(departments []DepartmentJSON) []FlatRecord {
	var out []FlatRecord
	for _, d := range departments {
		for _, g := range d.Groups {
			for _, s := range g.Students {
				for _, a := range s.Attendance {
					out = append(out, FlatRecord{
						Department: d.Department,
						Group:      g.Group,
						Student:    s.Student,
						Date:       a.Date,
						Missed:     a.Missed,
					})
				}
			}
		}
	}
	return out
}
