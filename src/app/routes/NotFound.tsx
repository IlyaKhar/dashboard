import { Button } from "@/shared/ui/button";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, FileQuestion } from "lucide-react";
import { twMerge } from "tailwind-merge";

export function NotFound() {
  const navigate = useNavigate();

  return (
    <div className="flex min-h-screen items-center justify-center bg-white">
      <div
        className={twMerge(
          "w-96 border border-gray-200 bg-white shadow-lg flex flex-col",
          "items-center gap-8 p-10 rounded-xl"
        )}
      >
        <div className="relative">
          <FileQuestion className="h-20 w-20 text-gray-400" />
        </div>

        <div className="text-center space-y-1.5">
          <div className="text-4xl font-medium text-black">404</div>
          <div className="text-lg text-gray-800">Страница не найдена</div>
        </div>

        <Button
          className="px-6 py-2.5 text-sm rounded-md"
          onClick={() => navigate("/")}
        >
          <span className="flex items-center gap-2">
            <ArrowLeft className="h-4 w-4" />
            На главную
          </span>
        </Button>
      </div>
    </div>
  );
}
