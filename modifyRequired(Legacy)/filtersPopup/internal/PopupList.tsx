import { twMerge } from "tailwind-merge";
import type { PopupListProps } from "../types";

export const PopupList = ({
  allSelectionText,
  noDataText,
  onPopupSelect,
  list,
  className,
}: PopupListProps) => {
  return (
    <ul className={twMerge("py-1 text-sm text-black", className)}>
      {allSelectionText && (
        <li
          onClick={() => onPopupSelect?.("")}
          className="cursor-pointer px-4 py-2 hover:bg-gray-100"
        >
          {allSelectionText}
        </li>
      )}
      {list.length ? (
        list.map((item) => (
          <li
            key={item}
            onClick={() => onPopupSelect?.(item)}
            className="cursor-pointer px-4 py-2 hover:bg-gray-100"
          >
            {item}
          </li>
        ))
      ) : (
        <div className="px-4 py-2 text-sm text-gray-400">{noDataText}</div>
      )}
    </ul>
  );
};
