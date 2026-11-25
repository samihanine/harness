import { chromium } from "playwright-core";
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
const p = await b.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
await p.goto(`http://localhost:3300${process.argv[2]}`);
await p.waitForTimeout(2500);
await p.screenshot({ path: ".local/crop.png", clip: { x: 0, y: 0, width: 1000, height: 260 } });
await b.close();
