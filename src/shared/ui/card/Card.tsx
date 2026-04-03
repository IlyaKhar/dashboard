import { twMerge } from "tailwind-merge";
import type { CardProps } from "./types";

export const Card = ({
  className,
  header,
  description,
  children,
  compact,
  ...props
}: CardProps) => (
  <div
    className={twMerge(
      "rounded-lg border bg-card text-card-foreground shadow-sm",
      className
    )}
    {...props}
  >
    <div
      className={twMerge(
        "flex flex-col space-y-1.5 p-6",
        compact ? "pb-6" : "pb-8"
      )}
    >
      <h3 className="text-2xl font-semibold leading-none tracking-tight">
        {header}
      </h3>
    </div>
    <div className={twMerge("p-6 pt-0", compact ? "pb-3 text-xl" : "")}>
      {children}
    </div>
    <div className="flex items-center p-6 pt-0 text-gray-600">
      {description}
    </div>
  </div>
);
