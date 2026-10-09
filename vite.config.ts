import { defineConfig, loadEnv } from "vite";
import { tanstackStart } from "@tanstack/react-start/plugin/vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig(({ mode }) => {
  // Networks that intercept TLS (corporate proxies): IGNORE_TLS_ERRORS=true in .env skips certificate
  // checks for the server's requests to Microsoft (sign-in, Power BI, Fabric, Graph).
  const env = loadEnv(mode, process.cwd(), "");
  if (env.IGNORE_TLS_ERRORS === "true") process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
  return {
    server: { port: 3300 },
    resolve: { tsconfigPaths: true },
    plugins: [tanstackStart({ spa: { enabled: true } }), react(), tailwindcss()],
  };
});
