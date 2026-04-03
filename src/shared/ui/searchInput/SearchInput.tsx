import { twMerge } from "tailwind-merge";
import { Search, X } from "lucide-react";

interface SearchInputProps {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  className?: string;
}

export const SearchInput = ({
  value,
  onChange,
  placeholder = "Поиск...",
  className,
}: SearchInputProps) => (
  <div className={twMerge("relative", className)}>
    <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-gray-400" />
    <input
      type="text"
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      className="h-9 w-full rounded-md border bg-white pl-9 pr-8 text-sm outline-none transition-colors placeholder:text-gray-400 focus:border-gray-400 focus:shadow-[0_0_0_1.5px_rgba(0,0,0,0.1)]"
    />
    {value && (
      <button
        type="button"
        onClick={() => onChange("")}
        className="absolute right-2 top-1/2 -translate-y-1/2 rounded-full p-0.5 text-gray-400 hover:text-gray-600"
        aria-label="Очистить поиск"
      >
        <X className="h-3.5 w-3.5" />
      </button>
    )}
  </div>
);
