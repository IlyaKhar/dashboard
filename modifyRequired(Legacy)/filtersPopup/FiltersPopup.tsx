import { useState } from "react";
import type { FiltersPopupProps } from "./types";
import { PopupList } from "./internal";
import { twMerge } from "tailwind-merge";

export const FiltersPopup = ({
  list,
  isOpen,
  width,
  maxHeight,
  searchInput,
  allSelectionText,
  noDataText,
  onChangeSearch,
  onPopupSelect,
  className,
}: FiltersPopupProps) => {
  const [overflow, setOverflow] = useState<"scroll" | "hidden">("hidden");

  return (
    <div
      style={{
        width,
        maxHeight: isOpen ? maxHeight : 0,
        overflowY: overflow,
      }}
      className={twMerge(
        "ml-2 absolute left-0 top-full mt-1.5 z-50 overflow-x-hidden ",
        "transition-[max-height,opacity,transform] duration-200 ease-out",
        "rounded-md border bg-white shadow-lg",
        className,
        isOpen
          ? "opacity-100 translate-y-0"
          : "opacity-0 -translate-y-1 pointer-events-none"
      )}
      onTransitionRun={() => setOverflow("hidden")}
      onTransitionEnd={() => isOpen && setOverflow("scroll")}
    >
      {searchInput && (
        <input
          className="w-full border-b px-3 py-2 text-sm focus:outline-none"
          placeholder="Поиск..."
          onChange={(e) => onChangeSearch?.(e.target.value)}
        />
      )}

      <PopupList
        onPopupSelect={onPopupSelect}
        list={list}
        allSelectionText={allSelectionText}
        noDataText={noDataText}
      />
    </div>
  );
};
