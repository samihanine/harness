import { chromium } from "playwright-core";
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
const p = await b.newPage({ viewport: { width: 900, height: 420 } });
const svg = await Bun.file(process.argv[2]).text();
await p.setContent(`<body style="margin:20px;background:#fff">${svg.replace("<svg ", '<svg width="800" ')}</body>`);
await p.screenshot({ path: ".local/svg.png" });
await b.close();
