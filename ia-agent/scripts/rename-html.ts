/** export-html: renames the single-file build to exports/<name>.html. */
import { readFileSync, renameSync, rmSync } from "node:fs";
import { join } from "node:path";

const root = join(import.meta.dir, "..");
const name = JSON.parse(readFileSync(join(root, "package.json"), "utf8")).name as string;
renameSync(join(root, "exports/html/index.html"), join(root, `exports/${name}.html`));
rmSync(join(root, "exports/html"), { recursive: true, force: true });
console.log(`exports/${name}.html`);
