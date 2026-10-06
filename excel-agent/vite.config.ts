import tailwindcss from "@tailwindcss/vite";
import { tanstackRouter } from "@tanstack/router-plugin/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";

/** SINGLE_FILE=1 builds one self-contained HTML (bun run export-html). */
const singleFile = process.env.SINGLE_FILE === "1";

export default defineConfig({
  base: "./",
  plugins: [
    tanstackRouter({ target: "react", autoCodeSplitting: !singleFile }),
    react(),
    tailwindcss(),
    ...(singleFile ? [viteSingleFile({ removeViteModuleLoader: true })] : []),
  ],
  resolve: { tsconfigPaths: true },
  server: { port: 3202, strictPort: true },
});
