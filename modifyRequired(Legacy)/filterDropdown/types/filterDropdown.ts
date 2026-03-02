export interface SizeData {
  width: number | string;
  height: number | string;
}

export interface FilterDropdownProps {
  label: string;
  list: string[];
  buttonSize?: SizeData;
  popupSize: SizeData;
  allSelection?: string;
  searchInput?: boolean;
  onChangeSearch?: (value: string) => void;
  onSelect?: (value: string) => void;
}