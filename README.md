# Rectorado UNJU — Sistema de gestión de tareas (prototipo)

Prototipo funcional con tres perfiles (Área, Administración, Dirección) para
el seguimiento de tareas, prioridades y avances del Rectorado.

⚠️ Los datos se guardan en el `localStorage` del navegador de cada persona
que usa la app — es decir, **no es una base de datos compartida en tiempo
real**. Cada usuario que entre desde su propio navegador va a ver los datos
de ejemplo iniciales y después los suyos propios. Para que todos vean la
misma información en vivo hace falta un backend con una base de datos real
(el siguiente paso natural del proyecto).

## Correr en tu computadora

1. Instalá [Node.js](https://nodejs.org/) (versión 18 o superior).
2. Abrí una terminal en esta carpeta y corré:
   ```bash
   npm install
   npm run dev
   ```
3. Abrí la URL que te muestra la terminal (normalmente `http://localhost:5173`).

## Subir el proyecto a GitHub

1. Creá un repositorio nuevo en GitHub (por ejemplo `rectorado-unju-tareas`), vacío, sin README.
2. En esta carpeta, corré:
   ```bash
   git init
   git add .
   git commit -m "Primera versión del sistema de gestión de tareas"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/rectorado-unju-tareas.git
   git push -u origin main
   ```

## Publicar la página online

### Opción A — GitHub Pages (gratis, usando el mismo repo)

1. En `vite.config.js`, cambiá la línea `base: "/"` por
   `base: "/rectorado-unju-tareas/"` (usando el nombre exacto de tu repositorio).
2. Instalá las dependencias si todavía no lo hiciste (`npm install`).
3. Corré:
   ```bash
   npm run deploy
   ```
   Esto genera la carpeta `dist` y la publica en una rama `gh-pages` del repo.
4. En GitHub, andá a **Settings → Pages** del repositorio y verificá que la
   fuente sea la rama `gh-pages`. La página va a quedar en:
   `https://TU_USUARIO.github.io/rectorado-unju-tareas/`

### Opción B — Netlify o Vercel (más simple, recomendada)

1. Subí el proyecto a GitHub (pasos de arriba).
2. Entrá a [netlify.com](https://netlify.com) o [vercel.com](https://vercel.com),
   conectá tu cuenta de GitHub e importá el repositorio.
3. Como comando de build usá `npm run build` y como carpeta de salida `dist`.
4. Listo — te da una URL pública en minutos y cada `git push` la actualiza sola.

## Próximo paso sugerido

Para que el sistema sea realmente compartido entre todas las áreas (no solo
en el navegador de cada uno), conviene sumar un backend simple (por ejemplo
Supabase o Firebase) que reemplace el `localStorage` por una base de datos
real con usuarios y contraseñas. Puedo ayudarte con ese paso cuando quieras
avanzar.
