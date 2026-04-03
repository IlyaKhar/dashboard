import type { HTMLAttributes } from "react";

export interface PieDiagramData {
  name: string;
  color: string;
  value: number;
}

export interface PieDiagramProps extends HTMLAttributes<HTMLDivElement> {
  data: PieDiagramData[];
  valueLabel?: string;
}
