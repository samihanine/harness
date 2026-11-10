/**
 * The slides of a PowerPoint deck drawn as pictures by LibreOffice itself, compiled to WebAssembly
 * (@matbee/libreoffice-converter: no installation, nothing leaves the machine). It runs beside the
 * local server, in its own process (office-engine.mjs), rather than in the page: no 240 MB download in
 * the browser, and the fonts installed on the computer are used (the ones the deck names are loaded
 * into the engine), so the picture is close to what PowerPoint shows, fonts included. One engine,
 * started at the first deck (or when the Presenter opens) and kept.
 */
import { readdir, readFile, stat } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";

/* ----------------------------------- Fonts ----------------------------------- */

const FONT_DIRS =
  process.platform === "win32"
    ? [join(process.env.WINDIR ?? "C:\\Windows", "Fonts"), join(process.env.LOCALAPPDATA ?? "", "Microsoft", "Windows", "Fonts")]
    : process.platform === "darwin"
      ? ["/System/Library/Fonts", "/Library/Fonts", join(homedir(), "Library/Fonts")]
      : ["/usr/share/fonts", "/usr/local/share/fonts", join(homedir(), ".fonts"), join(homedir(), ".local/share/fonts")];

/** The family names of a font file (its "name" table: ids 1 and 16), for .ttf / .otf and each font of a .ttc. */
export function fontFamilies(bytes: Uint8Array): string[] {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out = new Set<string>();
  const font = (start: number) => {
    const tables = v.getUint16(start + 4);
    for (let i = 0; i < tables; i++) {
      const rec = start + 12 + i * 16;
      if (String.fromCharCode(...bytes.subarray(rec, rec + 4)) !== "name") continue;
      const table = v.getUint32(rec + 8);
      const count = v.getUint16(table + 2);
      const strings = table + v.getUint16(table + 4);
      for (let n = 0; n < count; n++) {
        const r = table + 6 + n * 12;
        const [platform, id, length, offset] = [v.getUint16(r), v.getUint16(r + 6), v.getUint16(r + 8), v.getUint16(r + 10)];
        if (id !== 1 && id !== 16) continue;
        const raw = bytes.subarray(strings + offset, strings + offset + length);
        out.add(platform === 1 ? String.fromCharCode(...raw) : new TextDecoder("utf-16be").decode(raw));
      }
    }
  };
  try {
    if (String.fromCharCode(...bytes.subarray(0, 4)) === "ttcf") for (let i = 0; i < v.getUint32(8); i++) font(v.getUint32(12 + i * 4));
    else font(0);
  } catch {
    // Not a font this reads: ignored.
  }
  return [...out].filter(Boolean);
}

/** The fonts installed on this computer: family (lower case) → its files. Read once. */
let installed: Promise<Map<string, string[]>> | undefined;
function installedFonts() {
  return (installed ??= (async () => {
    const index = new Map<string, string[]>();
    const walk = async (dir: string, depth = 0): Promise<void> => {
      for (const e of await readdir(dir, { withFileTypes: true }).catch(() => [])) {
        const path = join(dir, e.name);
        if (e.isDirectory() && depth < 4) await walk(path, depth + 1);
        else if (/\.(ttf|otf|ttc)$/i.test(e.name)) for (const family of fontFamilies(await readFile(path).catch(() => new Uint8Array()))) index.set(family.toLowerCase(), [...(index.get(family.toLowerCase()) ?? []), path]);
      }
    };
    for (const dir of FONT_DIRS) if (dir) await walk(dir);
    return index;
  })());
}

/* ----------------------------------- Engine ---------------------------------- */

export type Rendered = { count: number; pages: { page: number; width: number; height: number; png: string }[] };
type Answer = Partial<Rendered> & { id?: number; ready?: boolean; error?: string };
type Engine = { process: ReturnType<typeof Bun.spawn>; fonts: Set<string>; ready: Promise<void>; waiting: Map<number, { resolve: (r: Rendered) => void; reject: (e: Error) => void }>; runtime: string };

const MAX_FONT_BYTES = 80_000_000;
const SCRIPT = join(import.meta.dir, "office-engine.mjs");
let engine: Engine | undefined;
let queue: Promise<unknown> = Promise.resolve();
let nextId = 1;

/** Node when the computer has it (the engine is far quicker to start there), else this Bun. */
const runtime = () => (process.env.HUB_OFFICE_RUNTIME === "bun" ? undefined : Bun.which("node")) ?? process.execPath;

/**
 * The engine's process (scripts/server/office-engine.mjs) holding the fonts asked for: started at the
 * first deck, started again when a deck names installed fonts it does not have yet.
 */
async function engineFor(families: string[]) {
  const index = await installedFonts();
  const wanted = families.map((f) => f.toLowerCase()).filter((f) => index.has(f));
  if (engine && engine.process.exitCode === null && wanted.every((f) => engine!.fonts.has(f))) return engine;
  const fonts = new Set([...(engine?.fonts ?? []), ...wanted]);
  const files: string[] = [];
  let size = 0;
  for (const family of fonts)
    for (const path of index.get(family) ?? []) {
      size += (await stat(path).catch(() => ({ size: 0 }))).size;
      if (size > MAX_FONT_BYTES) break;
      files.push(path);
    }
  // Not bundled with the server (240 MB): present after "bun install"; without it the app draws its own simple preview.
  if (!(await Bun.file(SCRIPT).exists()) || !(await Bun.file(join(import.meta.dir, "../../node_modules/@matbee/libreoffice-converter/package.json")).exists())) throw new Error("ENGINE_MISSING");
  engine?.process.kill();
  const waiting: Engine["waiting"] = new Map();
  let ready!: () => void;
  let failed!: (e: Error) => void;
  const started = new Promise<void>((resolve, reject) => ((ready = resolve), (failed = reject)));
  const child = Bun.spawn([runtime(), SCRIPT, ...files], {
    cwd: join(import.meta.dir, "../.."),
    stdio: ["ignore", "ignore", "ignore"],
    serialization: "json",
    ipc(message: Answer) {
      if (message.ready) return ready();
      const w = waiting.get(message.id ?? -1);
      if (!w) return;
      waiting.delete(message.id!);
      message.error ? w.reject(new Error(message.error)) : w.resolve({ count: message.count ?? 0, pages: message.pages ?? [] });
    },
    onExit() {
      const gone = new Error("The preview engine stopped.");
      failed(gone);
      for (const w of waiting.values()) w.reject(gone);
      waiting.clear();
      if (engine?.process === child) engine = undefined;
    },
  });
  engine = { process: child, fonts, ready: started, waiting, runtime: runtime() };
  return engine;
}

/* ----------------------------------- Render ---------------------------------- */

/**
 * Draws slides of a deck (`pages`: 1-based numbers; none: all) at a width in pixels. One deck at a
 * time (one engine). Throws "ENGINE_MISSING" when the engine is not installed.
 */
export function renderDeck(deck: Uint8Array, { width, pages, fonts }: { width: number; pages?: number[]; fonts: string[] }): Promise<Rendered> {
  const run = queue.then(async () => {
    const e = await engineFor(fonts);
    await e.ready;
    const id = nextId++;
    const answer = new Promise<Rendered>((resolve, reject) => e.waiting.set(id, { resolve, reject }));
    e.process.send({ id, deck: Buffer.from(deck).toString("base64"), width, pages });
    return answer;
  });
  queue = run.catch(() => undefined);
  return run;
}

/** Starts the engine ahead of the first deck (the Presenter asks when it opens): a tiny deck is not needed, the process only has to be up. */
export const warmEngine = () => void engineFor([]).catch(() => undefined);
