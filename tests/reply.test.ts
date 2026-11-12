import { describe, expect, test } from "bun:test";
import { parseReply } from "@/agent/reply";
import { layoutRows } from "@/features/builder/layout";
import { checkName } from "@/services/model";

describe("replies of a weak model", () => {
  test("a final answer whose quotes break the JSON is still read", () => {
    const reply = '```json{"tools": [{"name": "answer", "args": {"text": "J\'ai créé 4 visuels sur la page "Ventes Europe" :\\n- **Total**\\n- un "top 10" produits"}}]}```';
    expect(parseReply(reply)).toEqual([{ name: "answer", args: { text: 'J\'ai créé 4 visuels sur la page "Ventes Europe" :\n- **Total**\n- un "top 10" produits' } }]);
  });

  test("arguments written next to the tool's name, one of them called name", () => {
    expect(parseReply('{"tools": [{"name": "add_pivot", "name": "Ventes OTB Chine", "values": ["[Sales]"]}]}')).toEqual([{ name: "add_pivot", args: { name: "Ventes OTB Chine", values: ["[Sales]"] } }]);
  });

  test("a list closed too early keeps its calls", () => {
    const calls = parseReply('{"tools": [{"name": "get_visual_data", "args": {"visual": "a"}}], {"name": "get_visual_data", "args": {"visual": "b"}}]}');
    expect(calls[0]).toEqual({ name: "get_visual_data", args: { visual: "a" } });
  });

  test("broken JSON that is not a lone answer stays an error", () => {
    expect(() => parseReply('{"tools": [{"name": "run_dax", "args": {"query": "EVALUATE "x""}}]}')).toThrow();
  });
});

describe("tool arguments read tolerantly", () => {
  const info = { tables: [{ name: "Zone", columns: [{ name: "ZoneDescription" }], measures: [] }, { name: "Mesures", columns: [], measures: [{ name: "Stock Total" }] }] } as never;
  test("a whole reference where a column name is expected", () => {
    expect(checkName(info, "Zone", "'Zone'[ZoneDescription]")).toMatchObject({ table: "Zone", name: "ZoneDescription", kind: "column" });
    expect(checkName(info, "'Zone'", "zonedescription")).toMatchObject({ table: "Zone", name: "ZoneDescription" });
    expect(checkName(info, "Mesures", "[Stock Total]")).toMatchObject({ table: "Mesures", name: "Stock Total", kind: "measure" });
  });

  test("rows a little too high for the page are scaled, far too high refused", () => {
    const rows = [{ height: 110, items: [{ visual: "a" }] }, { height: 450, items: [{ visual: "b" }] }, { height: 450, items: [{ visual: "c" }] }];
    const boxes = layoutRows(rows, { width: 1280, height: 1080 });
    const last = boxes.at(-1)!;
    expect(last.y + last.height).toBeLessThanOrEqual(1080 - 20);
    expect(() => layoutRows(rows.map((r) => ({ ...r, height: r.height * 2 })), { width: 1280, height: 1080 })).toThrow(/rows take/);
  });
});

describe("instructions to the model", () => {
  test("never search the internet: first line of the protocol, and repeated at the end of every message", async () => {
    const { PROTOCOL, NO_WEB } = await import("@/prompts/agent");
    expect(PROTOCOL.startsWith("IMPORTANT — NEVER search the internet")).toBe(true);
    expect(NO_WEB).toMatch(/NEVER search the internet or the web/);
    const source = await Bun.file("src/agent/loop.ts").text();
    expect(source).toMatch(/\$\{NO_WEB\}\nNow:/);
  });
});
