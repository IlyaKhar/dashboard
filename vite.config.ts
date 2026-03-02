import react from "@vitejs/plugin-react-swc";
import { defineConfig } from "vite";
import path from "path";

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
  optimizeDeps: {
    include: ["react-is"],
  },
  server: {
    // IPv4 — чтобы localhost работал на Windows (без host браузер иногда не подключается)
    host: "127.0.0.1",
    port: 5173,
    // Прокси для API-запросов в dev-режиме.
    // Все запросы /api/* пойдут на Go-бэкенд.
    proxy: {
      "/api": {
        target:
          process.env.VITE_PROXY_TARGET || "http://localhost:8080",
        changeOrigin: true,
      },
    },
  },
});
