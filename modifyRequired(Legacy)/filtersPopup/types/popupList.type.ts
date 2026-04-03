import type { HTMLAttributes } from "react";

export interface PopupListProps extends HTMLAttributes<HTMLUListElement> {
  onPopupSelect?: (value: string) => void;
  list: string[];
  allSelectionText?: string;
  noDataText?: string;
}
