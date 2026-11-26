/**
 * A fake TrustLayer v2 (as documented) to check the Maia provider's requests and its reading of answers:
 * the first call answers a function call, the next ones a text. Run: bun scripts/checks/maia-mock.ts
 */
import { runChat } from "@server/agent/run";
import { settings } from "@server/settings";
let calls = 0;
const server = Bun.serve({
  port: 3399,
  async fetch(req) {
    const body = await req.json();
    calls++;
    const problems: string[] = [];
    if (new URL(req.url).pathname !== "/v2/chat/generation/") problems.push("path");
    if (req.headers.get("client_id") !== "id" || req.headers.get("client_secret") !== "secret") problems.push("headers");
    if (!["ANTHROPIC", "AZURE", "GEMINI", "OPENAI"].includes(body.provider)) problems.push("provider");
    if (!body.infos?.conversationId || !body.infos?.maison || !body.infos?.useCaseId || !body.infos?.callId || body.infos?.callStep === undefined || !body.infos?.environment) problems.push("infos");
    if (!body.tools?.[0]?.toolsDeclarations?.length) problems.push("tools");
    for (const m of body.messages) if (!["user", "model", "tool"].includes(m.role) || !Array.isArray(m.parts)) problems.push(`message ${m.role}`);
    console.log(`call ${calls}: roles ${body.messages.map((m: any) => m.role).join(",")} · ${body.tools?.[0]?.toolsDeclarations?.length} tools · ${problems.length ? `PROBLEMS ${problems}` : "ok"}`);
    if (calls === 1) return Response.json({ message: { role: "model", parts: [{ type: "functionCall", id: "call_1", name: "run_dax", parameters: { query: 'EVALUATE ROW("n", [Opportunity Count])' } }] }, usage: { inputTokens: 10, outputTokens: 5 } });
    const tool = body.messages.find((m: any) => m.role === "tool");
    console.log("tool result sent back:", JSON.stringify(tool?.parts?.[0]).slice(0, 160));
    if (calls === 2) return Response.json({ message: { role: "model", parts: [{ type: "functionCall", id: "call_2", name: "ask_user", parameters: { message: "458 opportunités.", questions: [{ question: "Détailler ?", options: ["Oui", "Non"] }] } }] } });
    return Response.json({ message: { role: "model", parts: [{ type: "text", text: "Il y a 458 opportunités." }] } });
  },
});
const s = await settings();
s.ai.provider = "maia";
s.ai.model = "claude-sonnet-5-5";
s.ai.maia = { ...s.ai.maia, url: "http://127.0.0.1:3399", clientId: "id", clientSecret: "secret" };
const scope = { dataset: { id: "eda77871-2401-4255-ba6c-2fc3c66902a1", name: "Revenue Opportunities" } };
for (const message of ["Combien d'opportunités ?", "Non"]) {
  let text = "";
  await runChat({ conversationId: "maia-check", screen: "trainer", message, scope: { report: { id: "2066ecc8-ead6-425b-a465-2475384528f1", name: "Revenue Opportunities", datasetId: scope.dataset.id } } }, (e) => {
    if (e.type === "text") text += e.text;
    if (e.type === "tool" && e.status !== "running") console.log(`  [${e.name} ${e.status}] ${(e.result ?? "").slice(0, 100)}`);
    if (e.type === "error") console.log("  ERROR", e.error);
  }, new AbortController().signal);
  console.log(`answer: ${text}`);
}
server.stop();
process.exit(0);
