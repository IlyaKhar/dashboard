import { useMemo } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "@/shared/ui/button";
import { Loader } from "@/shared/ui/loader";
import { ErrorState } from "@/shared/ui/errorState";
import { EmptyState } from "@/shared/ui/emptyState";
import { ArrowLeft } from "lucide-react";
import { useFetch } from "@/shared/hooks/useFetch";
import {
  fetchDrillDepartments,
  fetchAttendanceReconcileDay,
} from "@/shared/api";
import { AttendanceTable } from "@/widgets/attendanceTable";
import { drillToAttendanceRows } from "@/pages/groups/_lib/drilldownUtils";
import type { DeptDrillItem, AttendanceReconcileResponse } from "@/shared/api";
import type { AttendanceData } from "@/widgets/attendanceTable/types";
import { getOperationalDate } from "@/shared/hooks/useOperationalDate";

const POLLING_INTERVAL = 30_000;

export function DepartmentsPage() {
  const navigate = useNavigate();
  const operationalDate = getOperationalDate();

  const { data: departments, isLoading, error, refetch } = useFetch(
    (signal) => fetchDrillDepartments(signal),
    []
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

  // Формируем данные по парам для каждого отделения
  const departmentAttendanceData = useMemo(() => {
    if (!departments || !lessonsData || lessonsData.length !== 6) {
      return new Map<string, AttendanceData[]>();
    }

    const map = new Map<string, AttendanceData[]>();

    // Для каждого отделения собираем данные по парам
    for (const dept of departments) {
      const attendanceByLesson: AttendanceData[] = [];

      for (let lessonIndex = 0; lessonIndex < 6; lessonIndex++) {
        const lessonData = lessonsData[lessonIndex];
        const deptData = lessonData.byDepartment.find(
          (d) => d.department === dept.department
        );

        if (deptData) {
          attendanceByLesson.push({
            max: deptData.planned,
            total: deptData.present,
          });
        } else {
          // Если данных нет для этой пары, показываем NaN
          attendanceByLesson.push({
            max: 0,
            total: Number.NaN,
          });
        }
      }

      map.set(dept.department, attendanceByLesson);
    }

    return map;
  }, [departments, lessonsData]);

  function handleClick(dept: DeptDrillItem) {
    navigate("/groups", {
      state: { department: dept.department, departmentData: dept },
    });
  }

  return (
    <div className="space-y-6">
      <Button
        onClick={() => navigate("/")}
        variant="default"
        className="h-10 px-4 py-2"
      >
        <ArrowLeft className="mr-2 h-4 w-4" />
        Вернуться назад
      </Button>

      {isLoading && <Loader text="Загрузка отделений..." />}
      {error && <ErrorState message={error} onRetry={refetch} />}
      {lessonsLoading && !lessonsData && (
        <Loader text="Загрузка данных по парам..." />
      )}
      {lessonsError && !lessonsData && (
        <ErrorState message={lessonsError} onRetry={refetchLessons} />
      )}
      {!isLoading && !error && !departments?.length && (
        <EmptyState title="Нет отделений" />
      )}

      {departments && departments.length > 0 && (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {departments.map((dept) => {
            // Используем реальные данные по парам, если они есть, иначе fallback
            const attendanceData =
              departmentAttendanceData.get(dept.department) ||
              drillToAttendanceRows(dept.total, dept.absent);

            return (
              <div
                key={dept.department}
                role="button"
                tabIndex={0}
                onClick={() => handleClick(dept)}
                onKeyDown={(e) => e.key === "Enter" && handleClick(dept)}
                className="cursor-pointer animate-fade-in"
              >
                <AttendanceTable
                  header={`${dept.department}`}
                  attendance={attendanceData}
                />
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
