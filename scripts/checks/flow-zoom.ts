/** The report viewer on a tall page: it must take the frame's width and scroll (screenshot .local/zoom-*.png). */
import { openApp, shot } from "./browser";
const ID = process.argv[2] ?? "75cbbf42-6159-4023-9053-1485050722b6";
const PAGE = process.argv[3] ?? "Round4";
const { ctx, page } = await openApp();
await page.goto(`http://localhost:3300/trainer?item=${ID}`);
await page.waitForTimeout(14000);
const frame = page.frameLocator("iframe").first();
await frame.getByText(PAGE, { exact: true }).first().click({ timeout: 8000 }).catch(() => console.log("page tab not clicked"));
await page.waitForTimeout(6000);
await shot(page, "zoom-tall");
await ctx.close();
