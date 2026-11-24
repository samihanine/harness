/**
 * `bun run export:server`: the app built, then the server compiled to one executable with every file it
 * reads at run time embedded (server/static.gen.ts lists them; it is emptied again afterwards).
 */
import { readdirSync, statSync } from "node:fs";
import { join, relative } from "node:path";

const run = (cmd: string[]) => {
  const p = Bun.spawnSync(cmd, { stdio: ["inherit", "inherit", "inherit"] });
  if (p.exitCode) process.exit(p.exitCode);
};
run(["bunx", "vite", "build"]);
const walk = (dir: string): string[] => readdirSync(dir).flatMap((n) => (statSync(join(dir, n)).isDirectory() ? walk(join(dir, n)) : [join(dir, n)]));
const files = [...walk("dist"), ...walk("assets/template"), "node_modules/powerbi-client/dist/powerbi.min.js", "node_modules/pdfjs-dist/build/pdf.min.mjs", "node_modules/pdfjs-dist/build/pdf.worker.min.mjs"].map((f) => relative(".", f).replaceAll("\\", "/"));
const gen = `/** Written by scripts/build-server.ts for the executable; empty when running from the sources. */\n${files.map((f, i) => `import f${i} from "../${f}" with { type: "file" };`).join("\n")}\nexport const EMBEDDED: Record<string, string> = {\n${files.map((f, i) => `  ${JSON.stringify(f)}: f${i},`).join("\n")}\n};\n`;
const empty = await Bun.file("server/static.gen.ts").text();
await Bun.write("server/static.gen.ts", gen);
try {
  const target = process.argv[2] ? ["--target", process.argv[2]] : [];
  run(["bun", "build", "--compile", ...target, "server/main.ts", "--outfile", `dist-server/hub-server${process.argv[2]?.includes("windows") ? ".exe" : ""}`]);
} finally {
  await Bun.write("server/static.gen.ts", empty);
}
console.log(`Embedded ${files.length} files. Run dist-server/hub-server (it serves the app on http://127.0.0.1:3300).`);
