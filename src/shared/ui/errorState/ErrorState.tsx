import { twMerge } from "tailwind-merge";
import { AlertTriangle } from "lucide-react";
import { Button } from "@/shared/ui/button";

interface ErrorStateProps {
  className?: string;
  message?: string;
  onRetry?: () => void;
}

export const ErrorState = ({
  className,
  message = "Не удалось загрузить данные",
  onRetry,
}: ErrorStateProps) => (
  <div
    className={twMerge(
      "flex flex-col items-center justify-center gap-3 py-16 text-red-400",
      className
    )}
  >
    <AlertTriangle className="h-12 w-12" />
    <p className="text-sm text-gray-600">{message}</p>
    {onRetry && (
      <Button variant="outline" onClick={onRetry} className="mt-2 h-9 px-4">
        Попробовать снова
      </Button>
    )}
  </div>
);
