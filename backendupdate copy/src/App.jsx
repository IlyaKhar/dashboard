import { BrowserRouter, Routes, Route } from "react-router-dom";
import { Sidebar } from "@/components/Sidebar";
import { AttendancePage } from "@/pages/AttendancePage";
import { GroupPage } from "@/pages/GroupPage";
import "./App.css";

function App() {
  return (
    <BrowserRouter>
      <div className="flex h-screen bg-white">
        <Sidebar />
        <main className="flex-1 overflow-y-auto">
          <div className="border-b">
            <div className="px-8 py-6">
              <div className="space-y-2">
                <h1 className="text-3xl font-semibold tracking-tight text-black">
                  Мониторинг посещаемости
                </h1>
                <p className="text-sm text-muted-foreground">
                  Пропуски по неуважительной причине
                </p>
              </div>
            </div>
          </div>
          <div className="px-8 py-8">
            <Routes>
              <Route path="/" element={<AttendancePage />} />
              <Route path="/group/:groupName" element={<GroupPage />} />
            </Routes>
          </div>
          <div className="border-t px-8 py-6">
            <p className="text-center text-sm text-muted-foreground">
              ГАПОУ ОКЭИ • Мониторинг посещаемости • {new Date().getFullYear()}
            </p>
          </div>
        </main>
      </div>
    </BrowserRouter>
  );
}

export default App;
