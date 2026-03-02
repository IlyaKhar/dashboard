import { twMerge } from "tailwind-merge";
import { Calendar } from "lucide-react";

interface DateFilterProps {
  className?: string;
  date: string;
  onDateChange: (value: string) => void;
}

export const DateFilter = ({
  className,
  date,
  onDateChange,
}: DateFilterProps) => (
  <div
    className={twMerge(
      "flex items-center gap-3 rounded-lg border bg-white px-4 py-2.5",
      className
    )}
  >
    <Calendar className="h-4 w-4 text-gray-400" />
    <label className="flex items-center gap-2 text-sm text-gray-600">
      День
      <input
        type="date"
        value={date}
        onChange={(e) => onDateChange(e.target.value)}
        className="rounded-md border px-2 py-1 text-sm text-gray-800 outline-none focus:border-gray-400"
      />
    </label>
  </div>
);
