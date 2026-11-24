/**
 * Files the server reads at run time: the built app (dist/), the Power BI and pdf.js scripts of the headless
 * browser, the deck template's pictures. From the disk when running from the sources; embedded in the
 * executable by `bun run export:server` (scripts/build-server.ts writes server/static.gen.ts).
 */
import { existsSync } from "node:fs";
import { join } from "node:path";
import { EMBEDDED } from "./static.gen";

const ROOT = join(import.meta.dir, "..");

/** A file by its path from the project root (e.g. "dist/index.html", "node_modules/pdfjs-dist/build/pdf.min.mjs"). */
export function staticFile(path: string) {
  const embedded = EMBEDDED[path];
  if (embedded) return Bun.file(embedded);
  const p = join(ROOT, path);
  return existsSync(p) ? Bun.file(p) : undefined;
}

export const hasApp = () => !!staticFile("dist/index.html");
