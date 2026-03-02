import { twMerge } from "tailwind-merge";
import type { ButtonProps } from "./types";
import { buttonVariants } from "./_variants";

export const Button = ({
  className,
  variant = "default",
  type,
  ...props
}: ButtonProps) => {
  return (
    <button
      className={twMerge(
        "focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-2",
        "inline-flex items-center justify-center rounded-md text-sm font-medium",
        "transition-colors disabled:pointer-events-none disabled:opacity-50",
        buttonVariants[variant],
        className
      )}
      type={type}
      {...props}
    />
  );
};
