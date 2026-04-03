import type { BadgeProps } from "@/shared/ui/badge/types";
import type { AttendanceData } from "@/widgets/attendanceTable/types";

type BadgeVariant = NonNullable<BadgeProps["variant"]>;

export function drillToAttendanceRows(total: number, absent: number): AttendanceData[] {
  const present = total - absent;
  return [
    ...Array(4).fill(null).map(() => ({ max: total, total: present })),
    { max: total, total: Number.NaN },
    { max: total, total: Number.NaN },
  ];
}

export function getAttendanceBadge(
  total: number,
  absent: number
): { variant: BadgeVariant; text: string } {
  if (total === 0) return { variant: "secondary", text: "—" };
  const percent = Math.round(((total - absent) / total) * 100);
  if (percent >= 95) return { variant: "default", text: `${percent}%` };
  if (percent >= 90) return { variant: "outline", text: `${percent}%` };
  return { variant: "destructive", text: `${percent}%` };
}

export function getRowBgClass(total: number, absent: number): string {
  if (total === 0) return "";
  const percent = ((total - absent) / total) * 100;
  if (percent >= 95) return "bg-green-50";
  if (percent >= 90) return "bg-yellow-50";
  return "bg-red-50";
}
