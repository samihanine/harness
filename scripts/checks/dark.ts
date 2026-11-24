import { chromium } from "playwright-core";
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
const p = await b.newPage({ viewport: { width: 1440, height: 900 }, colorScheme: "dark" });
await p.goto(`http://localhost:3300${process.argv[2] ?? "/"}`);
await p.waitForTimeout(2500);
await p.screenshot({ path: ".local/dark.png" });
await b.close();
