import { captureReport, closeBrowser } from "@server/capture/browser";
const t = Date.now();
const r = await captureReport({ id: "2066ecc8-ead6-425b-a465-2475384528f1" });
console.log("first", Date.now() - t, r.page, r.visuals.length);
await Bun.write(process.argv[2] ?? ".local/capture.png", r.png);
const t2 = Date.now();
const r2 = await captureReport({ id: "2066ecc8-ead6-425b-a465-2475384528f1" }, { page: r.pages[1]?.name });
console.log("second", Date.now() - t2, r2.page);
await closeBrowser();
