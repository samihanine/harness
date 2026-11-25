/**
 * What the Power BI editor accepts through powerbi-report-authoring (editor mode): a report embedded in edit
 * mode in the headless browser, a list of operations tried on a scratch page (nothing saved), each result printed.
 * bun scripts/checks/authoring-probe.ts [reportId] [groupId] < ops.js  (ops: the body of an async function (page, report, log))
 */
import { chromium } from "playwright-core";
import { accessToken } from "@server/ms/auth";
const id = process.argv[2] ?? "75cbbf42-6159-4023-9053-1485050722b6";
const group = process.argv[3];
const ops = await Bun.stdin.text();
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium", args: ["--no-sandbox"], proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY } : undefined });
const page = await b.newPage({ viewport: { width: 1400, height: 900 } });
await page.setContent(`<!doctype html><html><body style="margin:0"><div id="r" style="width:1400px;height:900px"></div></body></html>`);
await page.addScriptTag({ path: "node_modules/powerbi-client/dist/powerbi.min.js" });
await page.addScriptTag({ path: "node_modules/powerbi-report-authoring/dist/powerbi-report-authoring.min.js" });
const token = await accessToken("powerbi");
page.on("console", (m) => console.log("[page]", m.text()));
const out = await page.evaluate(
  async ({ id, group, token, ops }) => {
    const pbi = (window as any)["powerbi-client"];
    const report = (window as any).powerbi.embed(document.getElementById("r"), { type: "report", id, embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${id}${group ? `&groupId=${group}` : ""}`, accessToken: token, tokenType: pbi.models.TokenType.Aad, permissions: pbi.models.Permissions.All, viewMode: ops.includes("//view") ? 0 : 1 });
    await new Promise((res, rej) => (report.on("rendered", res), report.on("error", (e: any) => rej(new Error(JSON.stringify(e.detail))))));
    const logs: string[] = [];
    const log = (...a: unknown[]) => logs.push(a.map((x) => (typeof x === "string" ? x : JSON.stringify(x))).join(" "));
    try {
      await new Function("report", "log", "pbi", `return (async () => { ${ops} })()`)(report, log, pbi);
    } catch (e: any) {
      log("THROWN", e?.message ?? JSON.stringify(e));
    }
    return logs;
  },
  { id, group, token, ops },
);
console.log(out.join("\n"));
await page.screenshot({ path: ".local/probe.png" });
await b.close();
