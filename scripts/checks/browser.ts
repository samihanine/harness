/** A browser for the UI checks, with its own profile (IndexedDB kept between runs: .local/pw-profile). */
import { chromium, type Page } from "playwright-core";

export async function openApp() {
  const ctx = await chromium.launchPersistentContext(".local/pw-profile", { executablePath: "/opt/pw-browsers/chromium", viewport: { width: 1440, height: 900 } }).catch(() => chromium.launchPersistentContext(".local/pw-profile", { viewport: { width: 1440, height: 900 } }));
  const page = ctx.pages()[0] ?? (await ctx.newPage());
  page.on("pageerror", (e) => console.log("pageerror", e.message));
  page.on("console", (m) => m.type() === "error" && !/404/.test(m.text()) && console.log("console", m.text().slice(0, 200)));
  return { ctx, page };
}

/** Sends a chat message and waits for the answer to finish (the Stop button gone). */
export async function ask(page: Page, text: string, timeout = 240_000) {
  await page.getByPlaceholder(/./).last().fill(text);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(1500);
  await page.getByRole("button", { name: "Stop" }).waitFor({ state: "detached", timeout });
}

/** The last assistant message's text and options. */
export async function lastAnswer(page: Page) {
  return page.evaluate(() => {
    const msgs = [...document.querySelectorAll(".markdown")];
    return msgs.slice(-3).map((m) => m.textContent).join("\n---\n");
  });
}

export async function shot(page: Page, name: string) {
  await page.screenshot({ path: `.local/${name}.png` });
}
