import { useMemo } from "react";
import { Card } from "@/shared/ui/card";
import { Loader } from "@/shared/ui/loader";
import { EmptyState } from "@/shared/ui/emptyState";
import { ErrorState } from "@/shared/ui/errorState";
import { PieDiagram } from "@/shared/ui/pieDiagram";
import { useFetch } from "@/shared/hooks/useFetch";
import { fetchDrillStudents } from "@/shared/api";
import type { DeptDrillItem, GroupDrillItem } from "@/shared/api";
import type { PieDiagramData } from "@/shared/ui/pieDiagram/types";

function getMissedHoursBadgeColor(missed: number): string {
  if (missed === 0) return "bg-[#16CF3E]";
  if (missed <= 4) return "bg-[#CF6316]";
  return "bg-[#CF1616]";
}

interface AnalyticsLevelProps {
  department: DeptDrillItem;
  group: GroupDrillItem;
}

export function AnalyticsLevel({ department, group }: AnalyticsLevelProps) {
  const {
    data: students,
    isLoading,
    error,
    refetch,
  } = useFetch(
    (signal) =>
      fetchDrillStudents(department.department, group.group, signal),
    [department.department, group.group]
  );

  const pieData = useMemo<PieDiagramData[]>(
    () => [
      { name: "Присутствует", value: group.total - group.absent, color: "#22c55e" },
      { name: "Отсутствует", value: group.absent, color: "#ef4444"},
    ],
    [group]
  );

  if (isLoading) return <Loader text="Загрузка аналитики..." />;
  if (error) return <ErrorState message={error} onRetry={refetch} />;

  return (
    <div className="space-y-6">
      <div className="animate-fade-in">
        <Card header={`Группа ${group.group}`}>
          <PieDiagram data={pieData} valueLabel="чел." />
        </Card>
      </div>

      {students && students.length > 0 && (
        <div className="animate-fade-in">
          <div
            className="overflow-x-auto rounded-[10px] border border-[rgba(123,123,123,0.5)] bg-white p-5"
              style={{ fontFamily: "Inter, system-ui, sans-serif" }}
            >
              <h2 className="mb-6 text-[24px] font-medium text-black">
                Список студентов группы {group.group}
              </h2>

              <div className="grid min-w-[700px] grid-cols-[50px_1fr_100px_minmax(120px,1fr)_140px] items-center gap-x-4 border-b border-[#C7C7C7] pb-3 text-left">
                <span className="text-[17px] font-medium">#</span>
                <span className="text-[16px] font-medium">ФИО студента</span>
                <span className="text-[16px] font-medium">Группа</span>
                <span className="text-[16px] font-medium">Отделение</span>
                <span className="text-right text-[16px] font-medium">Пропущено часов</span>
              </div>

              {students.map((s, idx) => (
                <div
                  key={s.student}
                  className="grid min-w-[700px] grid-cols-[50px_1fr_100px_minmax(120px,1fr)_140px] items-center gap-x-4 border-b border-[#C7C7C7] py-4 text-left"
                >
                  <div className="flex h-[27px] w-[27px] shrink-0 items-center justify-center rounded-full bg-black text-[14px] font-medium text-white">
                    {idx + 1}
                  </div>
                  <span className="text-[17px] font-medium text-black">{s.student}</span>
                  <span className="inline-flex h-5 w-fit min-w-[40px] items-center justify-center rounded-[37px] bg-[#F1F5F9] px-2.5 text-[14px] font-medium text-black">
                    {group.group}
                  </span>
                  <span
                    className="truncate text-[14px] font-medium text-[#8F8F8F]"
                    title={department.department}
                  >
                    {department.department}
                  </span>
                  <div className="flex justify-end">
                    <span
                      className={`inline-flex h-5 min-w-[45px] items-center justify-center rounded-[37px] px-2 text-[12px] font-medium text-white ${getMissedHoursBadgeColor(s.missed_total)}`}
                    >
                      {s.missed_total} ч.
                    </span>
                  </div>
                </div>
              ))}
            </div>
        </div>
      )}

      {students && students.length === 0 && (
        <EmptyState
          title="Нет студентов"
          description="В данной группе нет студентов в контингенте"
        />
      )}
    </div>
  );
}
