// Joins slide pictures into one sheet (for checks): bun scripts/checks/montage.ts out.png a.png b.png …
import { chromium } from "playwright-core";
const [out, ...files] = process.argv.slice(2);
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" }).catch(() => chromium.launch());
const p = await b.newPage({ viewport: { width: 1280, height: 200 } });
const imgs = await Promise.all(files.map(async (f) => `data:image/png;base64,${Buffer.from(await Bun.file(f).arrayBuffer()).toString("base64")}`));
await p.setContent(`<body style="margin:0;display:grid;grid-template-columns:1fr 1fr;gap:4px;background:#888">${imgs.map((s) => `<img src="${s}" style="width:100%">`).join("")}</body>`);
await p.screenshot({ path: out, fullPage: true });
await b.close();
