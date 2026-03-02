import type { SeparatorProps } from "./types";
import { separatorOrientations } from "./_variants";
import { twMerge } from "tailwind-merge";

export const Separator = ({
  className,
  orientation = "horizontal",
  decorative = true,
  ...props
}: SeparatorProps) => (
  <div
    role={decorative ? "none" : "separator"}
    aria-orientation={orientation}
    className={twMerge(
      "shrink-0 bg-slate-200",
      separatorOrientations[orientation],
      className
    )}
    {...props}
  />
); 

