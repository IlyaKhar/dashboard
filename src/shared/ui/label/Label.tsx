import { twMerge } from "tailwind-merge";
import type { LabelProps } from "./types";

export const Label = ({ className, ...props }: LabelProps) => (
  <label
    className={twMerge(
      "text-sm font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 text-black",
      className
    )}
    {...props}
  />
);
