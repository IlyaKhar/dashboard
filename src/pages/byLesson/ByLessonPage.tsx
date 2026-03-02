import { useState, useMemo } from "react";
import { Card } from "@/shared/ui/card";
import { Loader } from "@/shared/ui/loader";
import { ErrorState } from "@/shared/ui/errorState";
import { Button } from "@/shared/ui/button";
import { useFetch } from "@/shared/hooks/useFetch";
import {
  fetchAttendanceReconcileDay,
  fetchReconcileDayLessonGroup,
} from "@/shared/api";
import { pluralizeRu } from "@/shared/utils/pluralizeRu";
import { getOperationalDate } from "@/shared/hooks/useOperationalDate";

const POLLING_INTERVAL = 30_000;
const PERSONS = { one: "человек", few: "человека", many: "человек" };
const LESSON_LABELS: Record<number, string> = {
  1: "1-я пара",
  2: "2-я пара",
  3: "3-я пара",
  4: "4-я пара",
  5: "5-я пара",
  6: "6-я пара",
};

function formatDateForInput(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** Контент с key — при смене date/lesson ремаунтится, useFetch получает свежие данные */
function ByLessonContent({
  date,
  lesson,
}: {
  date: string;
  lesson: number;
}) {
  const [selectedGroup, setSelectedGroup] = useState<{
    group: string;
    department: string;
  } | null>(null);

  const {
    data: reconcile,
    isLoading: reconcileLoading,
    error: reconcileError,
    refetch: refetchReconcile,
  } = useFetch(
    (signal) => fetchAttendanceReconcileDay(date, signal, lesson),
    [date, lesson],
    { pollingInterval: POLLING_INTERVAL }
  );

  const {
    data: groupDetail,
    isLoading: groupDetailLoading,
    error: groupDetailError,
    refetch: refetchGroupDetail,
  } = useFetch(
    (signal) =>
      selectedGroup
        ? fetchReconcileDayLessonGroup(date, lesson, selectedGroup.group, signal)
        : Promise.resolve(null as never),
    [date, lesson, selectedGroup?.group ?? null]
  );

  const groupsFlat = useMemo(() => {
    if (!reconcile?.byDepartment) return [];
    const list: Array<{
      group: string;
      department: string;
      discipline: string;
      teacher: string;
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
          teacher: grp.teacher ?? "",
          planned: grp.planned,
          present: grp.present,
          absent: grp.absent,
        });
      }
    }
    return list;
  }, [reconcile]);

  if (reconcileLoading && !reconcile) return <Loader text="Загрузка..." />;
  if (reconcileError && !reconcile) {
    return (
      <ErrorState message={reconcileError} onRetry={refetchReconcile} />
    );
  }

  return (
    <div className="space-y-6">
      <p className="text-muted-foreground text-sm">
        Выберите номер пары — отобразятся группы и дисциплины за сегодня. Клик по
        группе: кто есть, кто нет.
      </p>

      {groupsFlat.length === 0 ? (
        <Card header={LESSON_LABELS[lesson]} compact>
          <p className="text-muted-foreground">
            На выбранную пару занятий не найдено.
          </p>
        </Card>
      ) : (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {groupsFlat.map((grp) => {
            const isSelected =
              selectedGroup?.group === grp.group &&
              selectedGroup?.department === grp.department;
            const percent =
              grp.planned > 0
                ? Math.round((grp.present / grp.planned) * 100)
                : 0;
            return (
              <Card
                key={`${grp.department}-${grp.group}`}
                header={grp.group}
                compact
                className={`cursor-pointer transition shadow-md ${
                  isSelected ? "ring-2 ring-black" : "hover:shadow-lg"
                }`}
                onClick={() =>
                  setSelectedGroup({
                    group: grp.group,
                    department: grp.department,
                  })
                }
              >
                <div className="space-y-1 text-sm">
                  {grp.discipline && (
                    <p className="font-medium text-foreground">
                      {grp.discipline}
                    </p>
                  )}
                  <p className="text-muted-foreground">
                    {grp.present} {pluralizeRu(grp.present, PERSONS)} из{" "}
                    {grp.planned} • {percent}%
                  </p>
                  <p className="text-red-600">
                    Нет: {grp.absent}
                  </p>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      {selectedGroup && (
        <Card
          header={`Группа ${selectedGroup.group} — ${LESSON_LABELS[lesson]}`}
          compact
        >
          {groupDetailLoading && !groupDetail ? (
            <Loader text="Загрузка состава..." />
          ) : groupDetailError && !groupDetail ? (
            <ErrorState
              message={groupDetailError}
              onRetry={refetchGroupDetail}
            />
          ) : groupDetail && "students" in groupDetail ? (
            <div className="space-y-2">
              {groupDetail.discipline && (
                <p className="font-medium">{groupDetail.discipline}</p>
              )}
              {groupDetail.teacher && (
                <p className="text-sm text-muted-foreground">
                  {groupDetail.teacher}
                </p>
              )}
              <p className="text-sm">
                Присутствует: {groupDetail.present} из {groupDetail.planned} • Нет:{" "}
                {groupDetail.absent}
              </p>
              <ul className="mt-3 max-h-64 overflow-y-auto rounded border divide-y divide-gray-200">
                {groupDetail.students.map((s) => (
                  <li
                    key={s.student}
                    className={`px-3 py-2 text-sm flex justify-between ${
                      s.present ? "bg-green-50" : "bg-red-50"
                    }`}
                  >
                    <span>{s.student}</span>
                    <span className={s.present ? "text-green-700" : "text-red-700"}>
                      {s.present ? "есть" : "нет"}
                    </span>
                  </li>
                ))}
              </ul>
              <Button
                variant="outline"
                className="mt-2"
                onClick={() => setSelectedGroup(null)}
              >
                Закрыть
              </Button>
            </div>
          ) : groupDetail && "message" in groupDetail ? (
            <p className="text-muted-foreground">
              {(groupDetail as { message?: string }).message}
            </p>
          ) : (
            <p className="text-muted-foreground">
              На эту пару у группы нет занятия.
            </p>
          )}
        </Card>
      )}
    </div>
  );
}

// Оперативный режим: дата = сегодня или дата с данными (16–22.02.2026)
export function ByLessonPage() {
  const operationalDate = getOperationalDate();
  const [lesson, setLesson] = useState(1);

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-center gap-4">
        <div className="flex flex-wrap gap-2">
          {([1, 2, 3, 4, 5, 6] as const).map((n) => (
            <Button
              key={n}
              variant={lesson === n ? "default" : "outline"}
              onClick={() => setLesson(n)}
            >
              {LESSON_LABELS[n]}
            </Button>
          ))}
        </div>
      </div>
      <ByLessonContent key={`${operationalDate}-${lesson}`} date={operationalDate} lesson={lesson} />
    </div>
  );
}
