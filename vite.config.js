import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],

  base: "/haia-xlsx-marge/", // הוסף כאן את שם המאגר שלך
});
