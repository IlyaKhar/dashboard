import { Card } from "@/shared/ui/card";
import type { TableData } from "@/shared/ui/table";
import Table from "@/shared/ui/table/Table";
import type { AttendanceTableProps } from "./types";

// Жёстко задаём сетку пар, чтобы можно было понять, какая пара сейчас идёт
// Время в минутах от начала суток
const PAIR_TIME_RANGES: Array<{ start: number; end: number }> = [
  { start: 8 * 60 + 30, end: 10 * 60 }, // 1 пара: 08:30–10:00
  { start: 10 * 60 + 10, end: 11 * 60 + 40 }, // 2 пара: 10:10–11:40
  { start: 12 * 60, end: 13 * 60 + 30 }, // 3 пара: 12:00–13:30
  { start: 14 * 60, end: 15 * 60 + 30 }, // 4 пара: 14:00–15:30
  { start: 15 * 60 + 40, end: 17 * 60 + 10 }, // 5 пара: 15:40–17:10
  { start: 17 * 60 + 15, end: 18 * 60 + 45 }, // 6 пара: 17:15–18:45
];

// Возвращает индекс текущей пары (0–5) или null, если сейчас ни одна пара не идёт
function getCurrentPairIndex(now: Date = new Date()): number {
  const minutes = now.getHours() * 60 + now.getMinutes();

  // До начала первой пары: ещё не началась ни одна
  if (minutes < PAIR_TIME_RANGES[0]!.start) return -1;

  for (let i = 0; i < PAIR_TIME_RANGES.length; i += 1) {
    const { start, end } = PAIR_TIME_RANGES[i]!;

    // Идёт i-я пара
    if (minutes >= start && minutes < end) {
      return i;
    }

    // Попали в «окно» между i-й и (i+1)-й парами:
    // считаем, что текущая — предыдущая (i), а все следующие ещё не начались.
    const next = PAIR_TIME_RANGES[i + 1];
    if (next && minutes >= end && minutes < next.start) {
      return i;
    }
  }

  // После последней пары: все пары уже прошли
  return PAIR_TIME_RANGES.length - 1;
}

function getAttendanceRowColor(
  present: number,
  max: number,
  thresholds: { green: number; yellow: number }
): string {
  if (max === 0 || Number.isNaN(present)) return "bg-gray-300";
  const percent = Math.round((present / max) * 100);
  if (percent >= thresholds.green) return "bg-green-300";
  if (percent >= thresholds.yellow) return "bg-yellow-300";
  return "bg-red-300";
}

export const AttendanceTable = ({
  attendance,
  colorSettings = { green: 80, yellow: 60 },
  header,
  onRowClick,
  selectedRowIndex,
  disableFutureHighlight,
}: AttendanceTableProps) => {
  // Определяем, какая пара идёт прямо сейчас (для сегодняшнего дня).
  // В историческом режиме (disableFutureHighlight) эта логика отключается.
  const currentPairIndex = getCurrentPairIndex();

  const tableData: TableData = {
    className: "text-base",
    header: {
      rows: [
        {
          cells: [
            {
              className: "p-1 w-0",
              type: "th",
              text: "Пара",
            },
            { type: "th", text: "Посещаемость" },
          ],
        },
      ],
    },
    body: {
      rows: attendance.map((el, i) => {
        // Если сейчас идёт N-я пара, то все следующие (i > N) считаем будущими и красим в серый.
        // В историческом режиме (disableFutureHighlight) не трогаем будущие пары вообще.
        const isFuturePair =
          !disableFutureHighlight && i > currentPairIndex;
        const rowBg = isFuturePair
          ? "bg-gray-200"
          : getAttendanceRowColor(el.total, el.max, colorSettings);

        const isSelected = selectedRowIndex === i;

        return {
          className: isSelected ? `${rowBg} ring-2 ring-black ring-inset` : rowBg,
          cells: [
            {
              className: `${rowBg} p-1 w-0 border-b-0`.trim(),
              type: "td",
              text: String(i + 1),
            },
            {
              type: "td",
              // Для будущих пар не показываем реальные цифры, а явно пишем, что пара ещё не началась.
              // В историческом режиме показываем фактические значения без этой приписки.
              text:
                isFuturePair && !disableFutureHighlight
                  ? "Пара ещё не началась"
                  : !Number.isNaN(el.total)
                    ? `${el.total} из ${el.max}`
                    : "---",
              className: `${rowBg} border-b-0`.trim(),
            },
          ],
          onClick: onRowClick ? () => onRowClick(i) : undefined,
        };
      }),
    },
  };

  return (
    <Card header={header}>
      <Table data={tableData} />
    </Card>
  );
};
