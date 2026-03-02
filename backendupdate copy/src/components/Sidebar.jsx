import { useState } from "react";
import { Link, useLocation } from "react-router-dom";
import { cn } from "@/lib/utils";
import { ClipboardList, ChevronLeft, ChevronRight } from "lucide-react";

export function Sidebar() {
  const location = useLocation();
  const [isCollapsed, setIsCollapsed] = useState(false);

  const navigation = [{ name: "Посещаемость", href: "/", icon: ClipboardList }];

  return (
    <div
      className={cn(
        "flex h-screen flex-col border-r bg-white transition-all duration-300",
        isCollapsed ? "w-16" : "w-64"
      )}
    >
      <div
        className={cn(
          "flex h-16 items-center border-b",
          isCollapsed ? "justify-center px-2" : "justify-between px-4"
        )}
      >
        {!isCollapsed && (
          <div className="flex items-center gap-3">
            <img
              src="/logo.png"
              alt="ГАПОУ ОКЭИ"
              className="h-8 w-8 flex-shrink-0"
            />
            <h1 className="text-xl font-semibold text-black">ГАПОУ ОКЭИ</h1>
          </div>
        )}
        <button
          onClick={() => setIsCollapsed(!isCollapsed)}
          className="rounded-lg p-2 text-gray-700 hover:bg-gray-100 transition-colors"
          title={isCollapsed ? "Развернуть" : "Свернуть"}
        >
          {isCollapsed ? (
            <ChevronRight className="h-5 w-5" />
          ) : (
            <ChevronLeft className="h-5 w-5" />
          )}
        </button>
      </div>
      <nav className="flex-1 space-y-1 px-3 py-4">
        {isCollapsed && (
          <div className="flex justify-center mb-4">
            <img src="/logo.png" alt="ГАПОУ ОКЭИ" className="h-8 w-8" />
          </div>
        )}
        {navigation.map((item) => {
          const isActive = location.pathname === item.href;
          return (
            <Link
              key={item.name}
              to={item.href}
              className={cn(
                "flex items-center gap-3 rounded-lg px-3 py-2 text-sm font-medium transition-colors",
                isActive
                  ? "bg-black text-white"
                  : "text-gray-700 hover:bg-gray-100",
                isCollapsed && "justify-center"
              )}
              title={isCollapsed ? item.name : undefined}
            >
              <item.icon className="h-5 w-5 flex-shrink-0" />
              {!isCollapsed && item.name}
            </Link>
          );
        })}
      </nav>
      {!isCollapsed && (
        <div className="border-t p-4">
          <p className="text-xs text-gray-600 text-center">
            Система мониторинга
          </p>
        </div>
      )}
    </div>
  );
}
