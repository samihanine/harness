import { openApp, shot } from "./browser";
const { ctx, page } = await openApp();
await page.goto(`http://localhost:3300${process.argv[2] ?? "/"}`);
await page.waitForTimeout(Number(process.argv[3] ?? 2500));
await shot(page, "peek");
await ctx.close();
