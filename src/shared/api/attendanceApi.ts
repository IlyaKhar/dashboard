import { apiRequest } from "./apiClient";
import type {
  AttendanceRecord,
  AttendanceSummaryResponse,
  AttendanceReconcileResponse,
  DeptDrillItem,
  GroupDrillItem,
  StudentDrillItem,
  LessonsDayResponse,
  DashboardStatsResponse,
  LoginResponse,
  LessonGroupDetailResponse,
} from "./types";

export interface LoginCredentials {
  username: string;
  password: string;
}

export function loginRequest(
  credentials: LoginCredentials
): Promise<LoginResponse> {
  return apiRequest<LoginResponse>("/login", {
    method: "POST",
    body: JSON.stringify(credentials),
  });
}

export interface AttendanceFilters {
  department?: string;
  group?: string;
  dateFrom?: string;
  dateTo?: string;
  period?: "7d" | "30d" | "90d";
}

function buildQueryString(filters: AttendanceFilters): string {
  const params = new URLSearchParams();
  if (filters.department) params.set("department", filters.department);
  if (filters.group) params.set("group", filters.group);
  if (filters.dateFrom) params.set("date_from", filters.dateFrom);
  if (filters.dateTo) params.set("date_to", filters.dateTo);
  if (filters.period) params.set("period", filters.period);
  const query = params.toString();
  return query ? `?${query}` : "";
}

export function fetchAttendance(
  filters: AttendanceFilters = {},
  signal?: AbortSignal
): Promise<AttendanceRecord[]> {
  return apiRequest<AttendanceRecord[]>(
    `/attendance${buildQueryString(filters)}`,
    { signal }
  );
}

export function fetchAttendanceSummary(
  filters: AttendanceFilters = {},
  signal?: AbortSignal
): Promise<AttendanceSummaryResponse> {
  return apiRequest<AttendanceSummaryResponse>(
    `/attendance/summary${buildQueryString(filters)}`,
    { signal }
  );
}

export function fetchAttendanceReconcileDay(
  date?: string,
  signal?: AbortSignal,
  lesson?: number
): Promise<AttendanceReconcileResponse> {
  const params = new URLSearchParams();
  if (date) params.set("date", date);
  if (lesson != null && lesson >= 1 && lesson <= 6) params.set("lesson", String(lesson));
  const query = params.toString();
  const suffix = query ? `?${query}` : "";
  return apiRequest<AttendanceReconcileResponse>(
    `/attendance/reconcile/day${suffix}`,
    { signal }
  );
}

export function fetchReconcileDayLessonGroup(
  date: string,
  lesson: number,
  group: string,
  signal?: AbortSignal
): Promise<LessonGroupDetailResponse> {
  const params = new URLSearchParams({ date, group });
  params.set("lesson", String(lesson));
  return apiRequest<LessonGroupDetailResponse>(
    `/attendance/reconcile/day/group?${params}`,
    { signal }
  );
}

export function fetchDrillDepartments(
  signal?: AbortSignal
): Promise<DeptDrillItem[]> {
  return apiRequest<DeptDrillItem[]>("/attendance/drill/departments", {
    signal,
  });
}

export function fetchDrillGroups(
  department: string,
  signal?: AbortSignal
): Promise<GroupDrillItem[]> {
  return apiRequest<GroupDrillItem[]>(
    `/attendance/drill/groups?department=${encodeURIComponent(department)}`,
    { signal }
  );
}

export function fetchDrillStudents(
  department: string,
  group: string,
  signal?: AbortSignal
): Promise<StudentDrillItem[]> {
  return apiRequest<StudentDrillItem[]>(
    `/attendance/drill/students?department=${encodeURIComponent(department)}&group=${encodeURIComponent(group)}`,
    { signal }
  );
}

export function fetchLessonsDay(
  date: string,
  signal?: AbortSignal
): Promise<LessonsDayResponse> {
  return apiRequest<LessonsDayResponse>(`/lessons/day?date=${date}`, {
    signal,
  });
}

export function fetchDashboardStats(
  signal?: AbortSignal
): Promise<DashboardStatsResponse> {
  return apiRequest<DashboardStatsResponse>("/dashboard/stats", { signal });
}
