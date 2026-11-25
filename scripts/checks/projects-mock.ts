/** A fake Confluence to check the projects' context: the pages of a folder given at the start of a conversation. */
import { settings } from "@server/settings";
import { projectContext } from "@server/projects";
const server = Bun.serve({
  port: 3398,
  fetch(req) {
    const u = new URL(req.url);
    if (u.pathname.endsWith("/content/search")) return Response.json({ results: [{ id: "11", title: "Glossaire" }, { id: "12", title: "Règles de calcul" }] });
    const id = u.pathname.match(/content\/(\d+)/)?.[1];
    const html = id === "11" ? "<h1>Glossaire</h1><p><strong>CA pondéré</strong> : revenu × probabilité.</p>" : "<p>Le CA se compte à la date de clôture estimée.</p>";
    return Response.json({ id, title: id === "11" ? "Glossaire" : "Règles de calcul", version: { number: 1 }, body: { storage: { value: html } } });
  },
});
const s = await settings();
s.confluence = { ...s.confluence, url: "http://127.0.0.1:3398/wiki", email: "x@y.z", key: "k", folderId: "999" };
console.log(await projectContext(["default"]));
server.stop();
