/** Production: serves the built front-end and the API on one port (`bun start`). */
import { join } from "node:path";
import { app } from "./app";

const PORT = Number(process.env.PORT ?? 3202);
const dist = join(import.meta.dir, "../dist");

Bun.serve({
  port: PORT,
  hostname: "127.0.0.1",
  async fetch(request) {
    const { pathname } = new URL(request.url);
    if (pathname.startsWith("/api/")) return app.fetch(request);
    const file = Bun.file(join(dist, pathname === "/" ? "index.html" : pathname));
    return (await file.exists()) ? new Response(file) : new Response(Bun.file(join(dist, "index.html")));
  },
});
console.log(`Excel agent on http://localhost:${PORT}`);
