/**
 * Round 5 UI (with scripts/checks/confluence-mock.ts and the app run on it): the catalog in the screens and in
 * Settings, the projects, the Data Scientist's workbook, the Presenter's tabs, the Builder's drop zone, Teams.
 * Screenshots .local/r5-*.png.
 */
import { openApp, shot } from "./browser";
const { ctx, page } = await openApp();
const go = async (path: string, wait = 2500) => (await page.goto(`http://localhost:3300${path}`), await page.waitForTimeout(wait));
await go("/");
await shot(page, "r5-home");
await go("/trainer?item=2066ecc8-ead6-425b-a465-2475384528f1", 12000);
await page.locator("h1").first().locator("xpath=..").locator("button").first().click();
await page.waitForTimeout(600);
await shot(page, "r5-trainer-picker");
await page.keyboard.press("Escape");
await page.getByRole("button", { name: /^Projects|No project/ }).first().click().catch(() => page.locator("header button.rounded-full").first().click());
await page.waitForTimeout(800);
await shot(page, "r5-projects-menu");
await page.keyboard.press("Escape");
await go("/settings?tab=confluence", 3000);
await shot(page, "r5-settings-confluence");
await go("/settings?tab=reports", 3000);
await shot(page, "r5-settings-reports");
await page.getByRole("button", { name: /^Add$/ }).click();
await page.getByPlaceholder(/paste the report's address/).fill("https://app.powerbi.com/groups/me/reports/75cbbf42-6159-4023-9053-1485050722b6/ReportSection");
await page.keyboard.press("Enter");
await page.waitForTimeout(6000);
await shot(page, "r5-settings-reports-added");
await go("/scientist", 2500);
const create = page.getByRole("button", { name: "New analysis" });
if (await create.count()) await create.first().click();
await page.keyboard.press("Escape");
await page.waitForTimeout(2500);
await page.getByRole("button", { name: /Semantic models/ }).first().click();
await page.waitForTimeout(1500);
await page.getByText("Revenue Opportunities").first().click();
await page.waitForTimeout(1500);
await page.getByRole("button", { name: /Workbook/ }).first().click();
await page.waitForTimeout(12000);
await shot(page, "r5-scientist");
await go("/presenter", 2500);
const newDeck = page.getByRole("button", { name: "New presentation" });
if (await newDeck.count()) await newDeck.first().click();
await page.keyboard.press("Escape");
await page.waitForTimeout(2500);
await shot(page, "r5-presenter-content");
await page.getByRole("button", { name: /Visual/ }).first().click();
await page.waitForTimeout(3000);
await shot(page, "r5-presenter-visual");
await go("/builder?item=", 2500);
await shot(page, "r5-builder");
await go("/teams", 6000);
await shot(page, "r5-teams");
await ctx.close();
