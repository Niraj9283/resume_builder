import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "node:path";

export default defineConfig({
  plugins: [react()],
  // Ensure Vite uses client/ as the project root so / and /index.html work
  root: path.resolve(__dirname, "."),
  server: {
    host: true,
    port: 5173,
    strictPort: true
  }
});
