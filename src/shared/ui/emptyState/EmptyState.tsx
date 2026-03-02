import { twMerge } from "tailwind-merge";
import { Inbox } from "lucide-react";

interface EmptyStateProps {
  className?: string;
  title?: string;
  description?: string;
}

export const EmptyState = ({
  className,
  title = "Нет данных",
  description = "Данные за выбранный период отсутствуют",
}: EmptyStateProps) => (
  <div
    className={twMerge(
      "flex flex-col items-center justify-center gap-3 py-16 text-gray-400",
      className
    )}
  >
    <Inbox className="h-12 w-12" />
    <div className="text-center">
      <p className="text-base font-medium text-gray-500">{title}</p>
      <p className="text-sm">{description}</p>
    </div>
  </div>
);
