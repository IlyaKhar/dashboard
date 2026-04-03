import type { HTMLAttributes } from "react";

export interface TableCell {
  type: "td" | "th";
  text: string;
  className?: string;
}

export interface TableRow {
  className?: string;
  cells: TableCell[];
  onClick?: () => void;
}

export interface TableHeader {
  className?: string;
  rows: TableRow[];
}

export interface TableBody {
  className?: string;
  rows: TableRow[];
}

export interface TableData {
  className?: string;
  header?: TableHeader;
  body?: TableBody;
}

export interface TableProps extends HTMLAttributes<HTMLDivElement> {
  data: TableData;
}
