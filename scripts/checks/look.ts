import { captureReport, closeBrowser } from "@server/capture/browser";
const r = await captureReport({ id: process.argv[2] }, { page: process.argv[3] });
await Bun.write(process.argv[4] ?? ".local/look.png", r.png);
await closeBrowser();
