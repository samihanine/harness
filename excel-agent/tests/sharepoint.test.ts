/**
 * SharePointTable against a simulated Graph Excel API (worksheets, tables, ranges, $batch):
 * checks the requests it sends and its row bookkeeping, without a Microsoft account.
 */
import { expect, test } from "bun:test";
import { SharePointTable } from "@/storage/sharepoint";

type Table = { name: string; header: string[]; rows: unknown[][] };

function fakeGraph() {
  const sheets = new Map<string, Table | null>();
  const log: string[] = [];
  let sessions = 0;
  const col = (n: number) => String.fromCharCode(65 + n);
  const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

  const handle = (method: string, url: string, body: any): { status: number; body: any } => {
    const path = decodeURIComponent(url.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, "").replace(/^\/drives\/d\/items\/f\/workbook/, ""));
    log.push(`${method} ${path}`);
    let m: RegExpExecArray | null;
    if (path === "/createSession") return { status: 201, body: { id: `s${++sessions}` } };
    if ((m = /^\/worksheets\/([^/?]+)\?/.exec(path))) return sheets.has(m[1]) ? { status: 200, body: { name: m[1] } } : { status: 404, body: { error: { code: "ItemNotFound", message: "no sheet" } } };
    if (path === "/worksheets/add") return sheets.set(body.name, null), { status: 201, body: { name: body.name } };
    if ((m = /^\/worksheets\/([^/]+)\/tables\?/.exec(path))) return { status: 200, body: { value: sheets.get(m[1]) ? [{ name: sheets.get(m[1])!.name }] : [] } };
    if ((m = /^\/worksheets\/([^/]+)\/usedRange/.exec(path))) return { status: 200, body: { address: `${m[1]}!A1:A1`, values: [[""]] } };
    if ((m = /^\/worksheets\/([^/]+)\/range\(address='([^']+)'\)$/.exec(path)) && method === "PATCH") {
      const sheet = m[1];
      const line = Number(/(\d+):/.exec(m[2])![1]);
      const t = sheets.get(sheet);
      const cells = (body.values ?? body.formulas)[0]; // rows are written as "formulas" (constants + formulas)
      if (!t) return sheets.set(sheet, { name: "", header: cells, rows: [] }), { status: 200, body: {} };
      if (line === 1) t.header = cells;
      else t.rows[line - 2] = cells;
      return { status: 200, body: {} };
    }
    if ((m = /^\/worksheets\/([^/]+)\/tables\/add$/.exec(path))) {
      const t = sheets.get(m[1])!;
      t.name = "Table1";
      t.rows = [t.header.map(() => "")]; // Excel adds one empty data row to a header-only table
      return { status: 201, body: { name: "Table1" } };
    }
    const table = [...sheets.entries()].find(([, t]) => t && path.includes(`/tables/${t.name}/`));
    if (table) {
      const [sheet, t] = table as [string, Table];
      if (path.endsWith("/range?$select=address,values") || path.endsWith("/range?$select=values"))
        return { status: 200, body: { address: `${sheet}!A1:${col(t.header.length - 1)}${t.rows.length + 1}`, values: [t.header, ...t.rows] } };
      if (path.endsWith("/columns/add")) {
        t.header.push(body.values[0][0]);
        t.rows.forEach((r) => r.push(""));
        return { status: 201, body: {} };
      }
      if (path.endsWith("/dataValidation")) return { status: 200, body: {} };
      if (path.endsWith("/rows/add")) return t.rows.push(...body.values), { status: 201, body: {} };
      if ((m = /rows\/\$\/itemAt\(index=(\d+)\)$/.exec(path)) && method === "DELETE") return t.rows.splice(Number(m[1]), 1), { status: 204, body: null };
    }
    return { status: 400, body: { error: { code: "Unhandled", message: `${method} ${path}` } } };
  };

  globalThis.fetch = (async (input: RequestInfo, init?: RequestInit) => {
    const url = String(input);
    if (url === "/api/token") return json({ token: "t", expiresAt: Date.now() + 3600_000 });
    const body = init?.body ? JSON.parse(String(init.body)) : undefined;
    if (url.endsWith("/$batch")) {
      const responses = body.requests.map((r: any) => {
        expect(r.headers["workbook-session-id"]).toBeTruthy();
        const out = handle(r.method, `https://graph.microsoft.com/v1.0${r.url}`, r.body);
        return { id: r.id, status: out.status, body: out.body };
      });
      log.push(`BATCH ${body.requests.length}`);
      return json({ responses });
    }
    const out = handle(init?.method ?? "GET", url, body);
    return out.status === 204 ? new Response(null, { status: 204 }) : json(out.body, out.status);
  }) as typeof fetch;
  return { sheets, log };
}

test("creates the sheet and table, writes in batches, keeps ids and indexes", async () => {
  const { sheets, log } = fakeGraph();
  const table = new SharePointTable({ driveId: "d", itemId: "f", name: "tasks.xlsx", webUrl: "" });
  const specs = [{ name: "id" }, { name: "title" }, { name: "due", date: true }, { name: "status", list: ["todo", "done"] }];
  expect(await table.open("Tasks", specs)).toEqual([]);
  const t = sheets.get("Tasks")!;
  expect(t.header).toEqual(["id", "title", "due", "status"]);

  table.append([
    { id: "a", title: "First", due: "2026-10-10", status: "todo" },
    { id: "b", title: "Second", due: null, status: "todo" },
    { id: "c", title: "Third", due: null, status: "done" },
  ]);
  table.update([{ id: "b", values: { status: "done" } }]);
  table.remove(["a"]);
  log.length = 0;
  await table.flush();
  // One round trip: fill the empty row, add 2 rows, update one row, delete one row.
  expect(log.filter((l) => l.startsWith("BATCH"))).toEqual(["BATCH 4"]);
  expect(t.rows).toEqual([
    ["b", "Second", "", "done"],
    ["c", "Third", "", "done"],
  ]);
  expect(await table.changedOutside()).toBe(false);

  // Someone edits the file: detected, and a reopen reads it (Excel date serial → ISO date).
  t.rows[1][2] = 46313; // 2026-10-18
  expect(await table.changedOutside()).toBe(true);
  const rows = await table.open("Tasks", specs);
  expect(rows.map((r) => [r.id, r.due])).toEqual([
    ["b", null],
    ["c", "2026-10-18"],
  ]);

  // A new schema column is added to the table; rows without id get one.
  t.rows.push(["", "No id", "", "todo"]);
  const again = await table.open("Tasks", [...specs, { name: "owner" }]);
  expect(t.header).toContain("owner");
  expect(again.every((r) => typeof r.id === "string" && r.id)).toBe(true);
  expect(t.rows[2][0]).toBe(again[2].id);
});
