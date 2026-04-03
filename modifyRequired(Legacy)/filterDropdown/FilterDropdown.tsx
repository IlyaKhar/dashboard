import { useRef, useState } from "react";
import { Button } from "@/shared/ui/button";
import { FiltersPopup } from "modifyRequired/filtersPopup";
import { LucideArrowDownWideNarrow } from "lucide-react";
import { useClickOutside } from "@/shared/hooks/useClickOutside";
import type { FilterDropdownProps } from "./types";

export const FilterDropdown = ({
  label,
  list,
  searchInput,
  popupSize,
  buttonSize,
  allSelection,
  onChangeSearch,
  onSelect,
}: FilterDropdownProps) => {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useClickOutside(rootRef, () => setOpen(false));

  return (
    <div ref={rootRef} className="relative inline-block">
      <Button style={buttonSize} onClick={() => setOpen((prev) => !prev)}>
        {label}
        <LucideArrowDownWideNarrow
          style={open ? { transform: "scaleY(-1)" } : {}}
          height={18}
          width={18}
          className="ml-1.5 transition-all duration-200 ease-out"
        />
      </Button>

      <FiltersPopup
        list={list}
        isOpen={open}
        maxHeight={popupSize?.height}
        width={popupSize?.width}
        onPopupSelect={(value) => {
          onSelect?.(value);
          setOpen(false);
        }}
        allSelectionText={allSelection}
        noDataText="Нет данных"
        onChangeSearch={(value) => onChangeSearch?.(value)}
        searchInput={searchInput}
      />
    </div>
  );
};
