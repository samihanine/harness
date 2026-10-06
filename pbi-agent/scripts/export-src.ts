/** Zips the app sources (without node_modules / builds) into exports/<name>-src.zip. */
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";
import { zipSync } from "fflate";

const root = join(import.meta.dir, "..");
const name = JSON.parse(readFileSync(join(root, "package.json"), "utf8")).name as string;
const skip = new Set(["node_modules", "dist", "exports", ".local", ".DS_Store", ".env"]);
const files: Record<string, Uint8Array> = {};

const walk = (dir: string) => {
  for (const entry of readdirSync(dir)) {
    if (skip.has(entry)) continue;
    const path = join(dir, entry);
    if (statSync(path).isDirectory()) walk(path);
    else files[`${name}/${relative(root, path)}`] = readFileSync(path);
  }
};
walk(root);
mkdirSync(join(root, "exports"), { recursive: true });
const out = join(root, "exports", `${name}-src.zip`);
writeFileSync(out, zipSync(files, { level: 9 }));
console.log(`${Object.keys(files).length} files → ${relative(root, out)}`);
