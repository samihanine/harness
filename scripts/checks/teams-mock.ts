/**
 * The Teams assistant on a fake chat (and a fake Confluence for its catalog): a data question (answered with a
 * table), a teasing (a funny answer with a GIF), people talking among themselves (no answer), its own message
 * never answered.  bun scripts/checks/teams-mock.ts
 */
import { settings } from "@server/settings";
import { useTeamsApi, type TeamsMessage } from "@server/teams/api";
import { teamsTick, updateTeams, saveGif } from "@server/teams/bot";
import { closeBrowser } from "@server/capture/browser";
const confluence = Bun.serve({
  port: 3397,
  fetch(req) {
    const u = new URL(req.url);
    if (u.pathname.endsWith("/api/v2/folders/801/direct-children")) return Response.json({ results: [{ id: "31", title: "Revenue Opportunities", type: "page" }] });
    if (u.pathname.endsWith("/content/31")) return Response.json({ id: "31", title: "Revenue Opportunities", version: { number: 1 }, body: { storage: { value: "<p>Power BI: https://app.powerbi.com/groups/me/datasets/eda77871-2401-4255-ba6c-2fc3c66902a1/details</p><p>Le pipeline commercial.</p>" } } });
    return Response.json({ results: [] });
  },
});
const s = await settings();
s.confluence = { ...s.confluence, url: "http://127.0.0.1:3397/wiki", email: "x@y.z", key: "k", datasetsFolderId: "801", reportsFolderId: "", projectsFolderId: "" };
await Bun.file(".local/library-dataset.json").delete().catch(() => undefined);
await Bun.file(".local/teams.json").delete().catch(() => undefined);
await saveGif("happy", new Uint8Array([71, 73, 70, 56, 57, 97]));
const messages: TeamsMessage[] = [];
const sent: string[] = [];
let n = 0;
const say = (who: string, text: string) => messages.push({ id: String(++n), createdDateTime: new Date(Date.now() + n * 1000).toISOString(), messageType: "message", from: { user: { id: who, displayName: who } }, body: { contentType: "html", content: `<p>${text}</p>` } });
useTeamsApi({
  me: async () => ({ id: "me", displayName: "Sami" }),
  chats: async () => [],
  lastMessage: async () => (messages.length ? { id: messages.at(-1)!.id, createdDateTime: messages.at(-1)!.createdDateTime } : undefined),
  messages: async (_c, top) => [...messages].reverse().slice(0, top),
  pictures: async () => [],
  async send(_c, html, gif) {
    sent.push(`${html}${gif ? " [GIF]" : ""}`);
    messages.push({ id: String(++n), createdDateTime: new Date(Date.now() + n * 1000).toISOString(), messageType: "message", from: { user: { id: "me", displayName: "Sami" } }, body: { contentType: "html", content: html } });
    return String(n);
  },
});
await updateTeams({ enabled: true, chatId: "chat-1", chatName: "Équipe ventes", name: "Scuba", datasetIds: ["eda77871-2401-4255-ba6c-2fc3c66902a1"], mode: "relevant" });
const turn = async (who: string, text: string) => {
  say(who, text);
  const before = sent.length;
  await teamsTick();
  console.log(`\n${who}: ${text}\n→ ${sent.length > before ? sent.at(-1) : "(no answer)"}`);
};
await turn("Alice", "Scuba, quel est le revenu par région ?");
await turn("Bob", "Quelqu'un vient au resto ce midi ?");
await turn("Alice", "t'es vraiment trop fort Scuba 😄 merci j'ai ma réponse");
await teamsTick();
console.log(`\n${sent.length} answers in all (its own messages never answered).`);
confluence.stop();
await closeBrowser();
