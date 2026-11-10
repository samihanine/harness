/**
 * The preview engine's own process: LibreOffice compiled to WebAssembly (@matbee/libreoffice-converter),
 * started by the local server (office.ts) with Node when the computer has it (the engine opens its
 * first deck in a second there, in more than a minute under Bun), else with Bun. It answers the
 * server's messages: a deck (base64) in, its slides as PNG (base64) out. Plain JavaScript, so both run it.
 */
import { readFileSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { createWorkerConverter } from "@matbee/libreoffice-converter/server";

/* ------------------------------------ PNG ------------------------------------ */

const CRC = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (bytes) => {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, "latin1");
  out.set(data, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
  return out;
};
/** The engine's pixels (BGRA) as a PNG. */
function png(pixels, width, height) {
  const rows = Buffer.alloc(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) {
    const to = y * (width * 4 + 1) + 1;
    const from = y * width * 4;
    for (let x = 0; x < width * 4; x += 4) {
      rows[to + x] = pixels[from + x + 2];
      rows[to + x + 1] = pixels[from + x + 1];
      rows[to + x + 2] = pixels[from + x];
      rows[to + x + 3] = 255;
    }
  }
  const head = Buffer.alloc(13);
  head.writeUInt32BE(width, 0);
  head.writeUInt32BE(height, 4);
  head.set([8, 6, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", head), chunk("IDAT", deflateSync(rows, { level: 3 })), chunk("IEND", Buffer.alloc(0))]);
}

/* ---------------------------------- Messages ---------------------------------- */

// The font files to load (the fonts of the computer the decks name), given as arguments.
const fonts = process.argv.slice(2).flatMap((path) => {
  try {
    return [{ filename: path.split(/[\\/]/).pop(), data: new Uint8Array(readFileSync(path)) }];
  } catch {
    return [];
  }
});
const converter = await createWorkerConverter({ fonts: fonts.length ? fonts : undefined, verbose: false });
const format = { inputFormat: "pptx" };

process.on("message", async ({ id, deck, width, pages }) => {
  try {
    const bytes = new Uint8Array(Buffer.from(deck, "base64"));
    const count = await converter.getPageCount(bytes, format);
    const indices = (pages?.length ? pages : Array.from({ length: count }, (_, i) => i + 1)).filter((p) => p >= 1 && p <= count).map((p) => p - 1);
    const drawn = indices.length ? await converter.renderPagePreviews(bytes, format, { width, pageIndices: indices }) : [];
    process.send({ id, count, pages: drawn.map((p) => ({ page: p.page + 1, width: p.width, height: p.height, png: png(p.data, p.width, p.height).toString("base64") })) });
  } catch (e) {
    process.send({ id, error: e instanceof Error ? e.message : String(e) });
  }
});
process.send({ ready: true });
