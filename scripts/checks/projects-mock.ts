/**
 * A fake Confluence to check the projects (sub-folders of the projects folder) and the catalog: a project's
 * pages given whole at the start of a conversation, read once and kept; the reports' catalog read from its pages.
 */
import { settings } from "@server/settings";
import { projectContext, projectsList } from "@server/projects";
import { library } from "@server/library";
let calls = 0;
const server = Bun.serve({
  port: 3398,
  fetch(req) {
    const u = new URL(req.url);
    calls++;
    if (u.pathname.endsWith("/api/v2/folders/900/direct-children")) return Response.json({ results: [{ id: "999", title: "Ventes", type: "folder" }, { id: "998", title: "RH", type: "folder" }] });
    if (u.pathname.endsWith("/api/v2/folders/800/direct-children")) return Response.json({ results: [{ id: "21", title: "Revenue Opportunities", type: "page" }] });
    if (u.pathname.includes("/api/v2/folders/")) return Response.json({ id: "999", _links: { webui: "/spaces/HUB/folder/999", base: "http://127.0.0.1:3398/wiki" } });
    if (u.pathname.endsWith("/content/search")) return Response.json({ results: [{ id: "11", title: "Glossaire" }, { id: "12", title: "Règles de calcul" }] });
    const id = u.pathname.match(/content\/(\d+)/)?.[1];
    if (id === "21") return Response.json({ id, title: "Revenue Opportunities", version: { number: 1 }, body: { storage: { value: "<p>Power BI: https://app.powerbi.com/groups/me/reports/2066ecc8-ead6-425b-a465-2475384528f1</p><p>Le pipeline commercial : revenus, opportunités, étapes.</p>" } } });
    const html = id === "11" ? "<h1>Glossaire</h1><p><strong>CA pondéré</strong> : revenu × probabilité.</p>" : "<p>Le CA se compte à la date de clôture estimée.</p>";
    return Response.json({ id, title: id === "11" ? "Glossaire" : "Règles de calcul", version: { number: 1 }, body: { storage: { value: html } } });
  },
});
const s = await settings();
s.confluence = { ...s.confluence, url: "http://127.0.0.1:3398/wiki", email: "x@y.z", key: "k", projectsFolderId: "900", reportsFolderId: "800", datasetsFolderId: "" };
await Bun.file(".local/projects/999.json").delete().catch(() => undefined);
await Bun.file(".local/library-report.json").delete().catch(() => undefined);
console.log("projects:", JSON.stringify(await projectsList()));
console.log(await projectContext(["999"]));
const first = calls;
await projectContext(["999"]);
console.log("calls first:", first, "second:", calls - first);
console.log("after reading:", JSON.stringify((await projectsList()).find((p) => p.id === "999")));
console.log("catalog:", JSON.stringify(await library("report", true)));
server.stop();
