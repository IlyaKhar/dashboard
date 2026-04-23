import { Link } from "react-router-dom";
import { useState, useMemo } from "react";
import { Card } from "@/shared/ui/card";
import { Loader } from "@/shared/ui/loader";
import { ErrorState } from "@/shared/ui/errorState";
import { AttendanceTable } from "@/widgets/attendanceTable";
import { useFetch } from "@/shared/hooks/useFetch";
import {
  fetchDashboardStats,
  fetchAttendanceSummary,
  fetchAttendanceReconcileDay,
} from "@/shared/api";
import { pluralizeRu } from "@/shared/utils/pluralizeRu";
import { drillToAttendanceRows } from "@/pages/groups/_lib/drilldownUtils";
import type { AttendanceData } from "@/widgets/attendanceTable/types";
import type { AttendanceReconcileResponse } from "@/shared/api";
import { getOperationalDate } from "@/shared/hooks/useOperationalDate";

const POLLING_INTERVAL = 30_000;
const PERSONS = { one: "человек", few: "человека", many: "человек" };

// Временные границы учебного дня (в минутах от начала суток)
const DAY_START_MINUTES = 8 * 60 + 30; // 08:30
const DAY_END_MINUTES = 18 * 60 + 45; // 18:45
const PAIR_TIME_RANGES: Array<{ start: number; end: number }> = [
  { start: 8 * 60 + 30, end: 10 * 60 }, // 1 пара
  { start: 10 * 60 + 10, end: 11 * 60 + 40 }, // 2 пара
  { start: 12 * 60, end: 13 * 60 + 30 }, // 3 пара
  { start: 14 * 60, end: 15 * 60 + 30 }, // 4 пара
  { start: 15 * 60 + 40, end: 17 * 60 + 10 }, // 5 пара
  { start: 17 * 60 + 15, end: 18 * 60 + 45 }, // 6 пара
];

function isLessonsTime(now: Date = new Date()): boolean {
  const minutes = now.getHours() * 60 + now.getMinutes();
  return minutes >= DAY_START_MINUTES && minutes <= DAY_END_MINUTES;
}

function getCurrentPairIndex(now: Date = new Date()): number {
  const minutes = now.getHours() * 60 + now.getMinutes();
  if (minutes < PAIR_TIME_RANGES[0]!.start) return -1;

  for (let i = 0; i < PAIR_TIME_RANGES.length; i += 1) {
    const { start, end } = PAIR_TIME_RANGES[i]!;
    if (minutes >= start && minutes < end) return i;
    const next = PAIR_TIME_RANGES[i + 1];
    if (next && minutes >= end && minutes < next.start) return i;
  }

  return PAIR_TIME_RANGES.length - 1;
}

// Оперативный режим: дата = сегодня, если в периоде демо-данных (16–22.02.2026), иначе 20.02.2026
export function HomePage() {
  const operationalDate = getOperationalDate();
  const [selectedLesson, setSelectedLesson] = useState<number | null>(null);

  const {
    data: stats,
    isLoading: statsLoading,
    error: statsError,
    refetch: refetchStats,
  } = useFetch(
    (signal) => fetchDashboardStats(signal),
    [],
    { pollingInterval: POLLING_INTERVAL }
  );

  // Оперативный: сверка за effectiveDate (сегодня или дата с данными)
  const {
    data: reconcile,
    isLoading: reconcileLoading,
    error: reconcileError,
    refetch: refetchReconcile,
  } = useFetch(
    (signal) => fetchAttendanceReconcileDay(operationalDate, signal),
    [operationalDate],
    { pollingInterval: POLLING_INTERVAL }
  );

  const {
    data: summary,
    isLoading: summaryLoading,
    error: summaryError,
    refetch: refetchSummary,
  } = useFetch(
    (signal) => fetchAttendanceSummary({}, signal),
    [],
    { pollingInterval: POLLING_INTERVAL }
  );

  // Оперативный: данные по парам за effectiveDate
  const {
    data: lessonsData,
    isLoading: lessonsLoading,
    error: lessonsError,
    refetch: refetchLessons,
  } = useFetch<AttendanceReconcileResponse[]>(
    (signal) =>
      Promise.all(
        ([1, 2, 3, 4, 5, 6] as const).map((n) =>
          fetchAttendanceReconcileDay(operationalDate, signal, n)
        )
      ),
    [operationalDate],
    { pollingInterval: POLLING_INTERVAL }
  );

  // Реальные данные по парам для обеих таблиц
  const attendanceByLesson = useMemo<AttendanceData[]>(() => {
    if (!lessonsData || lessonsData.length !== 6) {
      // Fallback на моковые данные, если реальные еще не загружены
      if (summary) {
        return drillToAttendanceRows(summary.total_students, summary.absent);
      }
      return Array(6)
        .fill(null)
        .map(() => ({ max: 0, total: Number.NaN }));
    }

    const hasAnyPlanned = lessonsData.some((r) => (r.totalPlanned ?? 0) > 0);
    if (!hasAnyPlanned) {
      // Если на выбранную дату расписание не подтянулось (все пары 0/0),
      // показываем согласованный fallback по summary, чтобы не было рассинхрона с карточками.
      if (summary) {
        return drillToAttendanceRows(summary.total_students, summary.absent);
      }
    }

    return lessonsData.map((r) => ({
      max: r.totalPlanned,
      total: r.totalPresent,
    }));
  }, [lessonsData, summary]);

  const isFirstLoad =
    (statsLoading && !stats) ||
    (summaryLoading && !summary) ||
    (reconcileLoading && !reconcile);
  const error = statsError || summaryError || reconcileError;

  if (isFirstLoad) return <Loader />;
  if (error && !stats && !summary) {
    return (
      <ErrorState
        message={error}
        onRetry={() => {
          refetchStats();
          refetchSummary();
          refetchReconcile();
        }}
      />
    );
  }

  const withinLessonsTime = isLessonsTime();
  const currentPairIndex = getCurrentPairIndex();
  const currentLessonReconcile =
    currentPairIndex >= 0 && currentPairIndex < 6
      ? lessonsData?.[currentPairIndex]
      : undefined;

  const hasCurrentLessonData = (currentLessonReconcile?.totalPlanned ?? 0) > 0;
  const hasDayReconcileData = (reconcile?.totalPlanned ?? 0) > 0;

  // Приоритет:
  // 1) текущая пара (если есть planned),
  // 2) агрегат по дню из reconcile (если есть planned),
  // 3) fallback на summary/stats (когда по расписанию на дату нет пар -> иначе UI "0/0").
  const presentCount = hasCurrentLessonData
    ? currentLessonReconcile!.totalPresent
    : hasDayReconcileData
      ? reconcile!.totalPresent
      : summary?.present ?? stats?.presentNow ?? 0;

  const absentCount = hasCurrentLessonData
    ? currentLessonReconcile!.totalAbsent
    : hasDayReconcileData
      ? reconcile!.totalAbsent
      : summary?.absent ?? stats?.absentNow ?? 0;

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <div className="grid grid-cols-1 gap-4">
          <Card
            header="Всего студентов присутствует"
            description={pluralizeRu(presentCount, PERSONS)}
            compact
          >
            {withinLessonsTime ? presentCount : "—"}
          </Card>
          <Card
            header="Всего студентов отсутствует"
            description={pluralizeRu(absentCount, PERSONS)}
            compact
          >
            {withinLessonsTime ? absentCount : "—"}
          </Card>
        </div>
        <div>
          <Link to="/departments">
            <AttendanceTable
              attendance={attendanceByLesson}
              header="Общая статистика по отделениям"
            />
          </Link>
        </div>
      </div>

      <div className="space-y-4 pt-4">
        <AttendanceTable
          attendance={attendanceByLesson}
          header="По парам — клик по строке: группы и посещаемость"
          onRowClick={(i) =>
            setSelectedLesson(selectedLesson === i + 1 ? null : i + 1)
          }
          selectedRowIndex={selectedLesson != null ? selectedLesson - 1 : undefined}
          colorSettings={{ green: 80, yellow: 60 }}
          disableFutureHighlight={false}
        />
        {lessonsLoading && !lessonsData && (
          <Loader text="Загрузка по парам..." />
        )}
        {lessonsError && !lessonsData && (
          <ErrorState message={lessonsError} onRetry={refetchLessons} />
        )}
        {selectedLesson != null && lessonsData?.[selectedLesson - 1] && (
          <GroupsByLessonCard
            reconcile={lessonsData[selectedLesson - 1]}
            lessonNumber={selectedLesson}
            onClose={() => setSelectedLesson(null)}
          />
        )}
      </div>
    </div>
  );
}

const LESSON_LABELS: Record<number, string> = {
  1: "1-я пара",
  2: "2-я пара",
  3: "3-я пара",
  4: "4-я пара",
  5: "5-я пара",
  6: "6-я пара",
};

function GroupsByLessonCard({
  reconcile,
  lessonNumber,
  onClose,
}: {
  reconcile: AttendanceReconcileResponse;
  lessonNumber: number;
  onClose: () => void;
}) {
  const groupsFlat = useMemo(() => {
    if (!reconcile?.byDepartment) return [];
    const list: Array<{
      group: string;
      department: string;
      discipline: string;
      planned: number;
      present: number;
      absent: number;
    }> = [];
    for (const dept of reconcile.byDepartment) {
      for (const grp of dept.byGroup) {
        list.push({
          group: grp.group,
          department: dept.department,
          discipline: grp.discipline ?? "",
          planned: grp.planned,
          present: grp.present,
          absent: grp.absent,
        });
      }
    }
    return list;
  }, [reconcile]);

  return (
    <Card
      header={`${LESSON_LABELS[lessonNumber]} — группы`}
      compact
      className="mt-2"
    >
      <div className="space-y-2">
        <p className="text-sm text-muted-foreground">
          Присутствует: {reconcile.totalPresent} из {reconcile.totalPlanned} •
          Нет: {reconcile.totalAbsent}
        </p>
        <ul className="max-h-64 overflow-y-auto rounded border divide-y divide-gray-200">
          {groupsFlat.map((grp) => {
            const percent =
              grp.planned > 0
                ? Math.round((grp.present / grp.planned) * 100)
                : 0;
            return (
              <li
                key={`${grp.department}-${grp.group}`}
                className="px-3 py-2 text-sm flex justify-between items-center"
              >
                <span>
                  {grp.group}
                  {grp.discipline && (
                    <span className="text-muted-foreground ml-2">
                      — {grp.discipline}
                    </span>
                  )}
                </span>
                <span>
                  {grp.present} из {grp.planned} ({percent}%)
                </span>
              </li>
            );
          })}
        </ul>
        <div className="mt-2 flex gap-4">
          <button
            type="button"
            onClick={onClose}
            className="text-sm text-muted-foreground hover:underline"
          >
            Закрыть
          </button>
          <Link
            to="/by-lesson"
            className="text-sm font-medium text-black hover:underline"
          >
            Подробнее: состав групп, кто есть/нет
          </Link>
        </div>
      </div>
    </Card>
  );
}
