/** After `vite build --mode html`: the one-file app copied to dist/hub.html (open it with the token server running). */
import { copyFile } from "node:fs/promises";
await copyFile("dist/html/index.html", "dist/hub.html");
console.log("dist/hub.html written (start the token server: bun run token, or dist/hub-token-server).");
