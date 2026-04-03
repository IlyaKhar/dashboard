import { useEffect } from "react";
import { BrowserRouter } from "react-router-dom";
import { AppRouter } from "./routes";
import { setApiErrorHandler } from "@/shared/api/apiErrorHandler";
import "./globals.css";

function App() {
  useEffect(() => {
    setApiErrorHandler((error, endpoint) => {
      console.error("[API Error]", endpoint, error.message);
    });
    return () => setApiErrorHandler(null);
  }, []);

  return (
    <BrowserRouter>
      <AppRouter />
    </BrowserRouter>
  );
}

export default App;
