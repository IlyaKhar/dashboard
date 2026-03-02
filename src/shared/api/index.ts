export { apiRequest } from "./apiClient";
export { setApiErrorHandler, notifyApiError } from "./apiErrorHandler";
export {
  loginRequest,
  fetchAttendance,
  fetchAttendanceSummary,
  fetchAttendanceReconcileDay,
  fetchReconcileDayLessonGroup,
  fetchDrillDepartments,
  fetchDrillGroups,
  fetchDrillStudents,
  fetchLessonsDay,
  fetchDashboardStats,
} from "./attendanceApi";
export type {
  LoginCredentials,
  AttendanceFilters,
} from "./attendanceApi";
export * from "./types";
