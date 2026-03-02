export interface LoginResponse {
  token: string;
  role: string;
}

export interface AttendanceRecord {
  department: string;
  group: string;
  student: string;
  date: string;
  missed: number;
}

export interface DepartmentSummary {
  department: string;
  total: number;
  absent: number;
  missed_total: number;
}

export interface AttendanceSummaryResponse {
  total_students: number;
  present: number;
  absent: number;
  by_department?: DepartmentSummary[];
}

export interface DeptDrillItem {
  department: string;
  total: number;
  absent: number;
  missed_total: number;
}

export interface GroupDrillItem {
  group: string;
  total: number;
  absent: number;
  missed_total: number;
}

export interface StudentDrillItem {
  student: string;
  missed_total: number;
  records: number;
  dates?: string[];
}

export interface LessonItem {
  group: string;
  department: string;
  discipline: string;
  dateTime: string;
  planned: number;
  present: number;
  percent: number;
}

export interface LessonsDayResponse {
  date: string;
  lessons: LessonItem[];
}

export interface DashboardStatsResponse {
  totalStudents: number;
  presentNow: number;
  absentNow: number;
  attendancePercent: number;
}

export interface ThresholdsResponse {
  id: number;
  type: string;
  green_threshold: number;
  yellow_threshold: number;
}

// Сверка расписания и посещаемости

export interface ReconcileGroupStats {
  group: string;
  planned: number;
  present: number;
  absent: number;
  discipline?: string;
  teacher?: string;
}

export interface LessonGroupStudent {
  student: string;
  present: boolean;
}

export interface LessonGroupDetailResponse {
  group: string;
  department: string;
  discipline: string;
  teacher: string;
  planned: number;
  present: number;
  absent: number;
  students: LessonGroupStudent[];
}

export interface ReconcileDepartmentStats {
  department: string;
  planned: number;
  present: number;
  absent: number;
  byGroup: ReconcileGroupStats[];
}

export interface AttendanceReconcileResponse {
  date: string;
  totalPlanned: number;
  totalPresent: number;
  totalAbsent: number;
  byDepartment: ReconcileDepartmentStats[];
}
