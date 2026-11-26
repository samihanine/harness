/**
 * A fake Confluence for UI checks (port 3396): a projects folder (900) with two projects, a reports folder (800)
 * and a models folder (801) with one page each; pages can be added, renamed and removed.
 *   bun scripts/checks/confluence-mock.ts   then run the app with CONFLUENCE_API_URL=http://127.0.0.1:3396/wiki
 *   CONFLUENCE_PROJECTS_FOLDER_ID=900 CONFLUENCE_REPORT_FOLDER_ID=800 CONFLUENCE_DATASET_FOLDER_ID=801 CONFLUENCE_CATALOG_EDIT=true
 */
type Page = { id: string; title: string; parent: string; html: string; version: number };
const pages = new Map<string, Page>();
const add = (id: string, parent: string, title: string, html: string) => pages.set(id, { id, title, parent, html, version: 1 });
const folders: Record<string, { id: string; title: string }[]> = { "900": [{ id: "910", title: "Ventes" }, { id: "920", title: "Finance" }] };
add("911", "910", "Glossaire des ventes", "<h1>Glossaire</h1><p><strong>CA pondéré</strong> : revenu × probabilité de l'étape.</p>");
add("912", "910", "Règles de calcul", "<p>Le CA se compte à la date de clôture estimée. Les montants sont hors taxes.</p>");
add("921", "920", "Budget", "<p>Le budget 2015 est de 2 Md$.</p>");
add("801", "800", "Revenue Opportunities", "<p>Power BI: https://app.powerbi.com/groups/me/reports/2066ecc8-ead6-425b-a465-2475384528f1</p><p>Le pipeline commercial 2015 : revenu, opportunités, étapes de vente, régions.</p>");
add("802", "801", "Revenue Opportunities", "<p>Power BI: https://app.powerbi.com/groups/me/datasets/eda77871-2401-4255-ba6c-2fc3c66902a1/details</p><p>Le modèle du pipeline commercial.</p>");
let next = 1000;
const page = (p: Page) => ({ id: p.id, title: p.title, version: { number: p.version }, body: { storage: { value: p.html } }, _links: { webui: `/pages/${p.id}`, base: "http://127.0.0.1:3396/wiki" } });
Bun.serve({
  port: 3396,
  async fetch(req) {
    const u = new URL(req.url);
    const path = u.pathname.replace(/^\/wiki/, "");
    let m = path.match(/^\/api\/v2\/folders\/(\d+)\/direct-children$/);
    if (m) {
      const id = m[1] === "801" ? "801-folder" : m[1];
      const list = [...(folders[m[1]] ?? []).map((f) => ({ ...f, type: "folder" })), ...[...pages.values()].filter((p) => p.parent === (m![1] === "801" ? "801" : m![1]) && !(m![1] === "801" && p.id === "801")).map((p) => ({ id: p.id, title: p.title, type: "page" }))];
      void id;
      return Response.json({ results: list });
    }
    m = path.match(/^\/api\/v2\/folders\/(\d+)$/);
    if (m) return Response.json({ id: m[1], spaceId: "1", _links: { webui: `/spaces/HUB/folder/${m[1]}`, base: "http://127.0.0.1:3396/wiki" } });
    if (path === "/api/v2/pages" && req.method === "POST") {
      const b = await req.json();
      const id = String(next++);
      add(id, b.parentId, b.title, b.body.value);
      return Response.json(page(pages.get(id)!));
    }
    if (path === "/rest/api/content/search") {
      const anc = decodeURIComponent(u.search).match(/ancestor=(\d+)/)?.[1];
      return Response.json({ results: [...pages.values()].filter((p) => p.parent === anc).map((p) => ({ id: p.id, title: p.title })) });
    }
    m = path.match(/^\/rest\/api\/content\/(\d+)$/);
    if (m) {
      const p = pages.get(m[1]);
      if (req.method === "DELETE") return pages.delete(m[1]), new Response(null, { status: 204 });
      if (req.method === "PUT" && p) {
        const b = await req.json();
        Object.assign(p, { title: b.title, html: b.body.storage.value, version: p.version + 1 });
      }
      return p ? Response.json(page(p)) : Response.json({ message: "Not found" }, { status: 404 });
    }
    return Response.json({ results: [] });
  },
});
console.log("Fake Confluence on http://127.0.0.1:3396/wiki");
