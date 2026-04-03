import type { HTMLAttributes } from "react";

export interface FiltersPopupProps extends HTMLAttributes<HTMLDivElement> {
  list: string[];
  isOpen: boolean;
  width: number | string;
  allSelectionText?: string;
  noDataText?: string;
  maxHeight: number | string;
  onChangeSearch?: (value: string) => void;
  onPopupSelect?: (value: string) => void;
  searchInput?: boolean;
}
