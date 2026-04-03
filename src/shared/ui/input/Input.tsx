import { twMerge } from "tailwind-merge";
import type { InputProps } from "./types";

export const Input = ({ className, type, ...props }: InputProps) => {
  return (
    <input
      type={type}
      className={twMerge(
        "flex h-10 w-full rounded-md border border-gray-300! bg-white px-3 py-2 text-sm",
        "file:border-0 file:bg-transparent file:text-sm file:font-medium",
        "placeholder:text-gray-500",
        "outline-none focus:shadow-[0_0_0_1.5px_rgba(0,0,0,0.7)]",
        "transition-all duration-200",
        "disabled:cursor-not-allowed disabled:opacity-50",
        className
      )}
      {...props}
    />
  );
};
