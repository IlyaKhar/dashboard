import { useState, useEffect } from "react";
import { Button } from "@/shared/ui/button";
import { LogIn, ArrowRight } from "lucide-react";
import { Outlet, useLocation, useNavigate } from "react-router-dom";

function formatDateTime(): { date: string; time: string } {
  const now = new Date();
  const formatted = now.toLocaleString("ru-RU", {
    day: "numeric",
    month: "long",
    hour: "2-digit",
    minute: "2-digit",
  });
  const [date, time] = formatted.split(" в ");
  return { date: date ?? "", time: time ?? "" };
}

export default function Layout() {
  const navigate = useNavigate();
  const location = useLocation();

  const [dateTime, setDateTime] = useState(formatDateTime);
  useEffect(() => {
    const interval = setInterval(() => setDateTime(formatDateTime()), 60_000);
    return () => clearInterval(interval);
  }, []);

  const isStoryPage = location.pathname === "/story";
  const historyButtonLabel = isStoryPage ? "Оперативный" : "Исторический";
  const historyButtonTarget = isStoryPage ? "/" : "/story";

  return (
    <div className="flex h-screen bg-white">
      <main className="flex flex-col flex-1 overflow-y-auto">
        <header className="border-b px-4 py-4 sm:px-8 sm:py-6 flex items-center justify-between gap-3">
          <h1 className="flex flex-col gap-1 sm:flex-row sm:gap-4.5 sm:items-end text-xl sm:text-3xl font-semibold tracking-tight">
            Мониторинг посещаемости
            <small className="text-gray-500 text-sm sm:text-lg flex items-end gap-2.5">
              {dateTime.date} <span>{dateTime.time}</span>
            </small>
          </h1>

          <Button
            onClick={() => navigate(historyButtonTarget)}
            className="bg-black hover:bg-gray-800 text-white min-h-11 rounded-full px-5 sm:px-6 flex items-center gap-2.5 text-base sm:text-lg shrink-0"
          >
            <span className="hidden sm:inline">{historyButtonLabel}</span>
            <span className="sm:hidden">{isStoryPage ? "Опер." : "Истор."}</span>
            <ArrowRight className="h-5 w-5" />
          </Button>
        </header>

        <div className="flex-1 p-3 pt-4 sm:p-4 sm:pt-5">
          <Outlet />
        </div>

        <footer className="mt-auto flex items-center justify-between border-t px-4 py-4 sm:px-8 sm:py-6 text-sm text-muted-foreground">
          <Button
            onClick={() => navigate("/logout")}
            className="bg-black hover:bg-gray-800 text-white min-h-11 px-5 text-base"
          >
            <LogIn className="mr-2.5 h-5 w-5" />
            Выйти
          </Button>
          <span className="text-xs sm:text-sm">
            ГАПОУ ОКЭИ • {new Date().getFullYear()}
          </span>
        </footer>
      </main>
    </div>
  );
}
