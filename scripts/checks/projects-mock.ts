/** A fake Confluence to check the projects' context: the pages of a folder given at the start of a conversation. */
import { settings } from "@server/settings";
import { projectContext, projectsStatus } from "@server/projects";
let calls = 0;
const server = Bun.serve({
  port: 3398,
  fetch(req) {
    const u = new URL(req.url);
    calls++;
    if (u.pathname.includes("/api/v2/folders/")) return Response.json({ id: "999", _links: { webui: "/spaces/HUB/folder/999", base: "http://127.0.0.1:3398/wiki" } });
    if (u.pathname.endsWith("/content/search")) return Response.json({ results: [{ id: "11", title: "Glossaire" }, { id: "12", title: "Règles de calcul" }] });
    const id = u.pathname.match(/content\/(\d+)/)?.[1];
    const html = id === "11" ? "<h1>Glossaire</h1><p><strong>CA pondéré</strong> : revenu × probabilité.</p>" : "<p>Le CA se compte à la date de clôture estimée.</p>";
    return Response.json({ id, title: id === "11" ? "Glossaire" : "Règles de calcul", version: { number: 1 }, body: { storage: { value: html } } });
  },
});
const s = await settings();
s.confluence = { ...s.confluence, url: "http://127.0.0.1:3398/wiki", email: "x@y.z", key: "k", folderId: "999" };
// Pages read once and kept (on disk): a second conversation reads nothing from Confluence; "Refresh" reads again.
await Bun.file(".local/projects/999.json").delete().catch(() => undefined);
console.log(await projectContext(["default"]));
const first = calls;
await projectContext(["default"]);
console.log("calls first:", first, "second:", calls - first, "status:", JSON.stringify(await projectsStatus()));
await projectsStatus("default");
console.log("after refresh:", calls - first, "calls");
server.stop();
