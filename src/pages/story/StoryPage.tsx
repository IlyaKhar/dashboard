import { useState, useMemo } from "react";
import { Card } from "@/shared/ui/card";
import { DateFilter } from "@/shared/ui/dateFilter";
import { Loader } from "@/shared/ui/loader";
import { EmptyState } from "@/shared/ui/emptyState";
import { ErrorState } from "@/shared/ui/errorState";
import { AttendanceTable } from "@/widgets/attendanceTable";
import { PieDiagram } from "@/shared/ui/pieDiagram";
import { useFetch } from "@/shared/hooks/useFetch";
import {
  fetchAttendanceReconcileDay,
  fetchAttendanceSummary,
} from "@/shared/api";
import type {
  AttendanceReconcileResponse,
  AttendanceSummaryResponse,
} from "@/shared/api";
import type { AttendanceData } from "@/widgets/attendanceTable/types";
import type { PieDiagramData } from "@/shared/ui/pieDiagram/types";

function formatDateForInput(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// Дефолтная дата истории — в периоде с данными (24–28.02), чтобы сводка сразу показывала актуальный период
const STORY_DEMO_START = "2026-02-24";
const STORY_DEMO_END = "2026-02-28";
const DEFAULT_STORY_DATE = (() => {
  const today = formatDateForInput(new Date());
  if (today >= STORY_DEMO_START && today <= STORY_DEMO_END) return today;
  return STORY_DEMO_START;
})();

/** Реальные данные по 6 парам из lessonsData; fallback — одна строка за день */
function storyDayToAttendanceRows(
  lessonsData: AttendanceReconcileResponse[] | null,
  reconcile: AttendanceReconcileResponse | null,
  summary: AttendanceSummaryResponse | null
): AttendanceData[] {
  if (lessonsData && lessonsData.length === 6) {
    return lessonsData.map((r) => ({
      max: r.totalPlanned,
      total: r.totalPresent,
    }));
  }
  if (reconcile?.totalPlanned && reconcile.totalPlanned > 0) {
    const max = reconcile.totalPlanned;
    const present = reconcile.totalPresent;
    return Array(6)
      .fill(null)
      .map(() => ({ max, total: present }));
  }
  if (summary?.total_students) {
    const max = summary.total_students;
    const present = summary.present;
    return Array(6)
      .fill(null)
      .map(() => ({ max, total: present }));
  }
  return Array(6)
    .fill(null)
    .map(() => ({ max: 0, total: Number.NaN }));
}

function buildPieData(
  reconcile: AttendanceReconcileResponse | null,
  summary: AttendanceSummaryResponse | null
): PieDiagramData[] {
  const colors = [
    "#4f46e5",
    "#06b6d4",
    "#f59e0b",
    "#ef4444",
    "#10b981",
    "#8b5cf6",
  ];

  if (reconcile?.byDepartment?.length) {
    return reconcile.byDepartment
      .filter((dept) => dept.department !== "Неизвестное отделение")
      .map((dept, i) => ({
        name: dept.department,
        value: dept.absent,
        color: colors[i % colors.length],
      }));
  }

  if (!summary?.by_department?.length) return [];

  return summary.by_department
    .filter((dept) => dept.department !== "Неизвестное отделение")
    .map((dept, i) => ({
      name: dept.department,
      value: dept.absent,
      color: colors[i % colors.length],
    }));
}

const POLLING_INTERVAL = 30_000;

export function StoryPage() {
  const [selectedDate, setSelectedDate] = useState(() => DEFAULT_STORY_DATE);

  // Сводка за день (для карточек и пирога)
  const {
    data: reconcile,
    isLoading: reconcileLoading,
    error: reconcileError,
    refetch: refetchReconcile,
  } = useFetch(
    (signal) => fetchAttendanceReconcileDay(selectedDate, signal),
    [selectedDate],
    { pollingInterval: POLLING_INTERVAL }
  );

  const {
    data: summary,
    isLoading: summaryLoading,
    error: summaryError,
    refetch: refetchSummary,
  } = useFetch(
    (signal) =>
      fetchAttendanceSummary(
        { dateFrom: selectedDate, dateTo: selectedDate },
        signal
      ),
    [selectedDate],
    { pollingInterval: POLLING_INTERVAL }
  );

  // Реальные данные по 6 парам для таблицы «Статистика посещаемости по дням»
  const {
    data: lessonsData,
    isLoading: lessonsLoading,
    error: lessonsError,
    refetch: refetchLessons,
  } = useFetch<AttendanceReconcileResponse[]>(
    (signal) =>
      Promise.all(
        ([1, 2, 3, 4, 5, 6] as const).map((n) =>
          fetchAttendanceReconcileDay(selectedDate, signal, n)
        )
      ),
    [selectedDate],
    { pollingInterval: POLLING_INTERVAL }
  );

  const attendance = useMemo<AttendanceData[]>(
    () =>
      storyDayToAttendanceRows(
        lessonsData ?? null,
        reconcile ?? null,
        summary ?? null
      ),
    [lessonsData, reconcile, summary]
  );

  const pieData = useMemo(
    () => buildPieData(reconcile, summary ?? null),
    [reconcile, summary]
  );

  const isLoading =
    reconcileLoading || summaryLoading || lessonsLoading;
  const error = reconcileError || summaryError || lessonsError;
  const hasData =
    (reconcile?.totalPlanned != null && reconcile.totalPlanned > 0) ||
    (summary?.total_students != null) ||
    (lessonsData != null && lessonsData.length === 6);
  const isEmpty = !isLoading && !error && !hasData;

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h2 className="text-xl font-semibold">Историческая сводка</h2>
        <DateFilter date={selectedDate} onDateChange={setSelectedDate} />
      </div>

      {isLoading && <Loader />}

      {error && (
        <ErrorState
          message={error}
          onRetry={() => {
            refetchReconcile();
            refetchSummary();
            refetchLessons();
          }}
        />
      )}

      {isEmpty && <EmptyState description="Нет данных за выбранный день" />}

      {!isLoading && !error && !isEmpty && (
        <>
          <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
            <div className="grid grid-cols-1 gap-4">
              <Card
                header="Всего студентов присутствует"
                description="человек"
                compact
              >
                {reconcile?.totalPlanned && reconcile.totalPlanned > 0
                  ? reconcile.totalPresent
                  : summary?.present ?? "—"}
              </Card>
              <Card
                header="Всего студентов отсутствует"
                description="человек"
                compact
              >
                {reconcile?.totalPlanned && reconcile.totalPlanned > 0
                  ? reconcile.totalAbsent
                  : summary != null
                  ? summary.absent ??
                    Math.max(
                      0,
                      (summary.total_students ?? 0) - (summary.present ?? 0)
                    )
                  : "—"}
              </Card>
            </div>
            <div>
              <AttendanceTable
                attendance={attendance}
                header="Статистика посещаемости по дням"
              />
            </div>
          </div>

          {pieData.length > 0 && (
            <Card header="Распределение пропусков по отделениям">
              <PieDiagram data={pieData} valueLabel="пропусков" />
            </Card>
          )}
        </>
      )}
    </div>
  );
}
