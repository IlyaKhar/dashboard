import { twMerge } from "tailwind-merge";

interface LoaderProps {
  className?: string;
  text?: string;
}

export const Loader = ({
  className,
  text = "Загрузка данных...",
}: LoaderProps) => (
  <div
    className={twMerge(
      "flex flex-col items-center justify-center gap-3 py-16",
      className
    )}
  >
    <div className="h-8 w-8 animate-spin rounded-full border-4 border-gray-200 border-t-gray-800" />
    <p className="text-sm text-gray-500">{text}</p>
  </div>
);
