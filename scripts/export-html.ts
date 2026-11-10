/**
 * After `vite build --mode html`: the one-file app copied to dist/hub.html. The local server serves it
 * (http://localhost:3301): opened there it has a real origin, which the server checks (docs/security.md).
 */
import { copyFile } from "node:fs/promises";
await copyFile("dist/html/index.html", "dist/hub.html");
console.log("dist/hub.html written. Start the local server (bun run token, or dist/hub-token-server beside hub.html) and open http://localhost:3301");
