import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANTE: si vas a publicar en GitHub Pages con un repo llamado, por
// ejemplo, "rectorado-unju-tareas", cambiá la línea de abajo a:
//   base: "/rectorado-unju-tareas/",
// Si vas a usar Netlify o Vercel, dejala en "/".
export default defineConfig({
  plugins: [react()],
  base: "/",
});
