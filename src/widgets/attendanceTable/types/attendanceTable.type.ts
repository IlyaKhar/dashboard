import type { CardProps } from "@/shared/ui/card";

export interface AttendanceData {
  max: number;
  total: number;
}

export interface ColorSettings {
  green: number;
  yellow: number;
}

export interface AttendanceTableProps extends CardProps {
  colorSettings?: ColorSettings;
  attendance: AttendanceData[];
  /** При клике по строке (индекс 0–5 = пара 1–6) */
  onRowClick?: (lessonIndex: number) => void;
  /** Индекс выбранной строки для подсветки */
  selectedRowIndex?: number;
}
