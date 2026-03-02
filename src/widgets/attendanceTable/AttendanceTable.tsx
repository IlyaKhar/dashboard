import { Card } from "@/shared/ui/card";
import type { TableData } from "@/shared/ui/table";
import Table from "@/shared/ui/table/Table";
import type { AttendanceTableProps } from "./types";

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
}: AttendanceTableProps) => {
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
        const rowBg = getAttendanceRowColor(el.total, el.max, colorSettings);
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
              text: !Number.isNaN(el.total) ? `${el.total} из ${el.max}` : "---",
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
