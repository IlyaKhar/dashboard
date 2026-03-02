import { twMerge } from "tailwind-merge";
import type { BadgeProps } from "./types";
import { badgeVariants } from "./_variants";

export const Badge = ({
  className,
  variant = "default",
  ...props
}: BadgeProps) => {
  return (
    <div
      className={twMerge(
        "focus:outline-none focus:ring-2 focus:ring-slate-950 focus:ring-offset-2",
        "rounded-full border px-2.5 py-0.5 text-xs font-semibold",
        "inline-flex items-center transition-colors",
        badgeVariants[variant],
        className
      )}
      {...props}
    />
  );
};
