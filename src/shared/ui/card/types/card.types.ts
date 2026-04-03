import type { HTMLAttributes } from "react";

export interface CardProps extends HTMLAttributes<HTMLDivElement> {
  header?: string;
  description?: string;
  compact?: boolean;
}
