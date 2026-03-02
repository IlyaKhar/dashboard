import { useMemo } from "react";
import { Loader } from "@/shared/ui/loader";
import { EmptyState } from "@/shared/ui/emptyState";
import { ErrorState } from "@/shared/ui/errorState";
import { AttendanceTable } from "@/widgets/attendanceTable";
import { useFetch } from "@/shared/hooks/useFetch";
import {
  fetchDrillGroups,
  fetchAttendanceReconcileDay,
} from "@/shared/api";
import { drillToAttendanceRows } from "../_lib/drilldownUtils";
import type { DeptDrillItem, GroupDrillItem, AttendanceReconcileResponse } from "@/shared/api";
import type { AttendanceData } from "@/widgets/attendanceTable/types";
import { getOperationalDate } from "@/shared/hooks/useOperationalDate";

const POLLING_INTERVAL = 30_000;

interface GroupsLevelProps {
  department: DeptDrillItem;
  onSelectGroup: (group: GroupDrillItem) => void;
}

export function GroupsLevel({ department, onSelectGroup }: GroupsLevelProps) {
  const operationalDate = getOperationalDate();

  const {
    data: groups,
    isLoading,
    error,
    refetch,
  } = useFetch(
    (signal) => fetchDrillGroups(department.department, signal),
    [department.department]
  );

  // Мониторинг: данные только за сегодня
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

  // Формируем данные по парам для каждой группы
  const groupAttendanceData = useMemo(() => {
    if (!groups || !lessonsData || lessonsData.length !== 6) {
      return new Map<string, AttendanceData[]>();
    }

    const map = new Map<string, AttendanceData[]>();

    // Для каждой группы собираем данные по парам
    for (const grp of groups) {
      const attendanceByLesson: AttendanceData[] = [];

      for (let lessonIndex = 0; lessonIndex < 6; lessonIndex++) {
        const lessonData = lessonsData[lessonIndex];
        // Ищем группу в отделении
        const deptData = lessonData.byDepartment.find(
          (d) => d.department === department.department
        );
        const groupData = deptData?.byGroup.find(
          (g) => g.group === grp.group
        );

        if (groupData) {
          attendanceByLesson.push({
            max: groupData.planned,
            total: groupData.present,
          });
        } else {
          // Если данных нет для этой пары, показываем NaN
          attendanceByLesson.push({
            max: 0,
            total: Number.NaN,
          });
        }
      }

      map.set(grp.group, attendanceByLesson);
    }

    return map;
  }, [groups, lessonsData, department.department]);

  if (isLoading) return <Loader text="Загрузка групп..." />;
  if (error) return <ErrorState message={error} onRetry={refetch} />;
  if (!groups?.length)
    return (
      <EmptyState
        title="Нет групп"
        description="Группы в данном отделении не найдены"
      />
    );

  return (
    <div className="space-y-4">
      {lessonsLoading && !lessonsData && (
        <Loader text="Загрузка данных по парам..." />
      )}
      {lessonsError && !lessonsData && (
        <ErrorState message={lessonsError} onRetry={refetchLessons} />
      )}
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3 animate-fade-in">
        {(groups ?? []).map((grp) => {
          // Используем реальные данные по парам, если они есть, иначе fallback
          const attendanceData =
            groupAttendanceData.get(grp.group) ||
            drillToAttendanceRows(grp.total, grp.absent);

          return (
            <div
              key={grp.group}
              role="button"
              tabIndex={0}
              onClick={() => onSelectGroup(grp)}
              onKeyDown={(e) => e.key === "Enter" && onSelectGroup(grp)}
              className="cursor-pointer"
            >
              <AttendanceTable
                header={grp.group}
                attendance={attendanceData}
              />
            </div>
          );
        })}
      </div>
    </div>
  );
}
