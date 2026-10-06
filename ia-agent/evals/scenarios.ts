/**
 * Realistic host apps for evaluating the harness: the same kinds of tools / resources as
 * excel-agent and pbi-agent expose, on small fixed data, with checks on the outcome.
 */
import { McpServer, ResourceTemplate } from "@modelcontextprotocol/server";
import { z } from "zod";

export type Outcome = { answer?: string; question?: { text: string; options: string[] } };
export type Scenario = {
  name: string;
  prompt: string;
  history?: string;
  server: () => { server: McpServer; check: (outcome: Outcome) => string[] };
};

const text = (value: unknown, undo?: { name: string; arguments: Record<string, unknown> }) => ({
  content: [{ type: "text" as const, text: typeof value === "string" ? value : JSON.stringify(value) }],
  ...(undo ? { _meta: { undo } } : {}),
});
const fail = (message: string) => ({ content: [{ type: "text" as const, text: message }], isError: true });
const page = (server: McpServer, description: string) =>
  server.registerResource("page", "app://page", { title: "Current page", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
    contents: [{ uri: uri.href, text: description }],
  }));
const newServer = (name: string, instructions: string) => new McpServer({ name, title: name, version: "1" }, { instructions });

/* ------------------------------- excel-agent ------------------------------ */

const EXCEL_INSTRUCTIONS = `The user edits local Excel files through schemas (field definitions).
The resource app://page says which page is open and what can be done there.
Values must match the field types; options must be one of the allowed values.`;

function schemaEditor(): ReturnType<Scenario["server"]> {
  const server = newServer("excel-agent", EXCEL_INSTRUCTIONS);
  let schema: any = { id: "s1", name: "Tasks", sheet: "Tasks", fields: [{ name: "title", type: "string", label: "Title", required: true }] };
  page(server, `Schema editor for "Tasks". The schema (resource schema://current) defines the columns of Excel files: field name (column header), type, label, description, required, and for "option" fields the allowed values (with a color) and "multiple". Change it with patch_schema (JSON Patch on the schema object).`);
  server.registerResource("schema", "schema://current", { title: "Schema being edited", mimeType: "application/json", annotations: { priority: 1 } }, async (uri) => ({
    contents: [{ uri: uri.href, text: JSON.stringify(schema, null, 1) }],
  }));
  server.registerTool(
    "patch_schema",
    {
      description:
        'Changes the schema with JSON Patch operations (RFC 6902) applied in order, e.g. {"op":"add","path":"/fields/-","value":{"name":"price","type":"number"}} or {"op":"replace","path":"/fields/2/label","value":"Price"}. Field types: string, text, number, integer, boolean, date, option, image. Option colors: gray, red, orange, amber, green, teal, blue, violet, pink. The id cannot change.',
      inputSchema: z.object({ operations: z.array(z.object({ op: z.enum(["add", "remove", "replace", "move", "copy", "test"]), path: z.string(), value: z.unknown().optional(), from: z.string().optional() })) }),
    },
    async ({ operations }) => {
      const next = structuredClone(schema);
      for (const op of operations) {
        const parts = op.path.split("/").slice(1);
        if (op.op === "add" && parts[0] === "fields") {
          if (parts[1] === "-") next.fields.push(op.value);
          else if (parts.length === 2) next.fields.splice(Number(parts[1]), 0, op.value);
          else next.fields[Number(parts[1])][parts[2]] = op.value;
        } else if (op.op === "replace" && parts[0] === "fields") next.fields[Number(parts[1])][parts[2]] = op.value;
        else if (op.op === "remove" && parts[0] === "fields") next.fields.splice(Number(parts[1]), 1);
        else if (op.op === "replace" && parts.length === 1) next[parts[0]] = op.value;
        else return fail(`Unsupported operation ${JSON.stringify(op)}`);
      }
      for (const f of next.fields) {
        if (!/^[A-Za-z_]\w*$/.test(f.name ?? "")) return fail(`fields: invalid name "${f.name}" (letters, digits and _ only)`);
        if (!["string", "text", "number", "integer", "boolean", "date", "option", "image"].includes(f.type)) return fail(`fields.${f.name}.type: invalid "${f.type}"`);
        if (f.type === "option" && !(f.options?.length && f.options.every((o: any) => typeof o.value === "string")))
          return fail(`fields.${f.name}.options: expected [{"value": "...", "color": "..."}]`);
      }
      schema = next;
      return text(`Schema updated (${schema.fields.length} fields).`);
    },
  );
  return {
    server,
    check: (outcome) => {
      const status = schema.fields.find((f: any) => f.type === "option" && /status|statut/i.test(f.name + (f.label ?? "")));
      return [!status && "no status option field", !outcome.answer && "no answer"].filter(Boolean) as string[];
    },
  };
}

type Task = { id: string; title: string; owner: string; status: string; due: string | null };
const TASKS: Task[] = [
  { id: "t1", title: "Prepare budget", owner: "Bob", status: "doing", due: "2026-10-10" },
  { id: "t2", title: "Review contract", owner: "Alice", status: "blocked", due: null },
  { id: "t3", title: "Call supplier", owner: "Bob", status: "todo", due: "2026-10-08" },
  { id: "t4", title: "Update website", owner: "Chloé", status: "blocked", due: "2026-10-20" },
  { id: "t5", title: "Plan offsite", owner: "Alice", status: "todo", due: null },
  { id: "t6", title: "Fix invoice 42", owner: "Bob", status: "blocked", due: "2026-10-09" },
];
const TASK_FIELDS = [
  { name: "title", type: "string", required: true },
  { name: "owner", type: "string" },
  { name: "status", type: "option", options: ["todo", "doing", "blocked", "done"].map((value) => ({ value })) },
  { name: "due", type: "date" },
];

function fileEditor(mode: "read" | "edit", expect: "count" | "bob" | "ask" | "followup" = mode === "read" ? "count" : "bob"): ReturnType<Scenario["server"]> {
  const server = newServer("excel-agent", EXCEL_INSTRUCTIONS);
  let rows = structuredClone(TASKS).map((r) => (expect === "followup" && r.owner === "Bob" ? { ...r, status: "done" } : r));
  const calls: string[] = [];
  page(
    server,
    `File editor: "Team tasks" (local file tasks.xlsx, schema "Tasks").
${rows.length} rows in total; the user's current view shows ${rows.length} rows (no filter).
excel://view holds the rows of the current view, excel://rows all rows (one JSON row per line, with its id).
${mode === "edit" ? "Edit mode: you may add, update and delete rows." : "Read mode: you can only read; if the user asks for changes, tell them to switch the agent to Edit mode (pencil in the toolbar)."}`,
  );
  server.registerResource("schema", "excel://schema", { title: "Fields of the rows", mimeType: "application/json", annotations: { priority: 1 } }, async (uri) => ({ contents: [{ uri: uri.href, text: JSON.stringify(TASK_FIELDS) }] }));
  server.registerResource("view", "excel://view", { title: "Rows of the current view", mimeType: "application/jsonl", annotations: { priority: 1 } }, async (uri) => ({
    contents: [{ uri: uri.href, text: rows.map((r) => JSON.stringify(r)).join("\n") }],
  }));
  if (mode === "edit") {
    const values = z.record(z.string(), z.unknown());
    const checkValues = (v: Record<string, unknown>) => {
      for (const [k, x] of Object.entries(v)) {
        if (k === "id") continue;
        if (!TASK_FIELDS.some((f) => f.name === k)) return `Unknown field "${k}" (fields: ${TASK_FIELDS.map((f) => f.name).join(", ")})`;
        if (k === "status" && !["todo", "doing", "blocked", "done"].includes(String(x))) return `status: unknown option(s) ${x} (allowed: todo, doing, blocked, done)`;
      }
      return "";
    };
    server.registerTool("add_rows", { description: "Adds rows at the end of the file.", inputSchema: z.object({ rows: z.array(values) }) }, async ({ rows: added }) => {
      calls.push("add_rows");
      for (const v of added) if (checkValues(v)) return fail(checkValues(v));
      const created = added.map((v, i) => ({ id: `n${rows.length + i + 1}`, title: "", owner: "", status: "todo", due: null, ...v }) as Task);
      rows = [...rows, ...created];
      return text({ addedIds: created.map((r) => r.id) });
    });
    server.registerTool(
      "update_rows",
      { description: "Changes some values of existing rows (only the given fields change).", inputSchema: z.object({ updates: z.array(z.object({ id: z.string(), values })) }) },
      async ({ updates }) => {
        calls.push("update_rows");
        for (const u of updates) {
          if (!rows.some((r) => r.id === u.id)) return fail(`No row with id ${u.id}`);
          if (checkValues(u.values)) return fail(checkValues(u.values));
        }
        rows = rows.map((r) => ({ ...r, ...(updates.find((u) => u.id === r.id)?.values ?? {}) }));
        return text({ updatedIds: updates.map((u) => u.id) });
      },
    );
    server.registerTool("delete_rows", { description: "Deletes rows by id.", inputSchema: z.object({ ids: z.array(z.string()) }) }, async ({ ids }) => {
      calls.push("delete_rows");
      rows = rows.filter((r) => !ids.includes(r.id));
      return text({ deletedIds: ids });
    });
  }
  return {
    server,
    check: (outcome) => {
      const problems: string[] = [];
      if (!outcome.answer) problems.push("no answer");
      const unchanged = (ids: string[]) => ids.every((id) => JSON.stringify(rows.find((r) => r.id === id)) === JSON.stringify(TASKS.find((t) => t.id === id)));
      if (expect === "ask") {
        problems.length = 0;
        if (!outcome.question) problems.push("did not ask before a destructive, ambiguous change");
        if (calls.length) problems.push(`changed data: ${calls.join(", ")}`);
        return problems;
      }
      if (!outcome.answer) problems.push("no answer");
      if (expect === "count") {
        if (!/\b3\b|three|trois/i.test(outcome.answer ?? "")) problems.push(`wrong count in answer: ${outcome.answer?.slice(0, 120)}`);
        if (calls.length) problems.push("changed data in read mode");
      } else if (expect === "followup") {
        if (rows.find((r) => r.id === "t3")?.status !== "todo") problems.push("t3 not back to todo");
        if (!unchanged(["t2", "t4", "t5"]) || rows.find((r) => r.id === "t1")?.status !== "done") problems.push("other rows changed");
      } else {
        const bob = rows.filter((r) => r.owner === "Bob");
        if (!bob.every((r) => r.status === "done")) problems.push(`Bob's tasks not all done: ${bob.map((r) => r.status)}`);
        if (!unchanged(["t2", "t4", "t5"])) problems.push("other rows changed");
        const added = rows.find((r) => /test/i.test(r.title));
        if (!added || added.owner !== "Alice") problems.push(`new task wrong: ${JSON.stringify(added)}`);
      }
      return problems;
    },
  };
}

/* -------------------------------- pbi-agent ------------------------------- */

const PBI_INSTRUCTIONS = `The user works with Power BI: a local database of reports, semantic models (datasets) and guides, a DAX runner, a researcher (pivot tables over several sources), a report builder and a report viewer.
The resource app://page says which page is open and what can be done there.
Remote datasets are queried in DAX (EVALUATE …); local datasets (Excel) in SQL with [brackets]. Never invent table, column or measure names: use the model resources.`;

const SALES_MODEL = `'Sales'
  columns: OrderDate:DateTime, ProductKey:Int64, CustomerKey:Int64, Quantity:Int64, Amount:Double
  measures: [Total Sales] = SUM(Sales[Amount]) | [Orders] = DISTINCTCOUNT(Sales[OrderID])
'Product'
  columns: ProductKey:Int64, Product:String, Category:String
'Date'
  columns: Date:DateTime, Year:Int64, Month:String
relationships:
  'Sales'[ProductKey] → 'Product'[ProductKey]
  'Sales'[OrderDate] → 'Date'[Date]`;

function daxRunner(): ReturnType<Scenario["server"]> {
  const server = newServer("pbi-agent", PBI_INSTRUCTIONS);
  let editor = "EVALUATE\nTOPN(100, 'Sales')";
  const queries: string[] = [];
  page(server, `Query runner on dataset "Contoso Sales" (Power BI semantic model: DAX, EVALUATE …). dax://model describes the model, dax://editor holds the user's query and its last result. Use run_query to test queries, set_editor to put a query in the user's editor.`);
  server.registerResource("model", "dax://model", { title: "Model of Contoso Sales", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({ contents: [{ uri: uri.href, text: SALES_MODEL }] }));
  server.registerResource("editor", "dax://editor", { title: "Editor and last result", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
    contents: [{ uri: uri.href, text: `Query (DAX):\n${editor}\n\nLast result: (not run)` }],
  }));
  const fakeRun = (q: string) => {
    queries.push(q);
    if (!/^\s*(DEFINE[\s\S]*)?EVALUATE/i.test(q)) return fail("Query must start with EVALUATE (or DEFINE … EVALUATE).");
    if (/\bProducts?\b'?\[(Name|ProductName)\]/i.test(q)) return fail("Column 'Name' cannot be found in table 'Product'.");
    if (!/Product'?\[Product\]/i.test(q)) return text("4 rows\nCategory,Total Sales\nBikes,120000\nAccessories,30000\nClothing,22000\nComponents,9000");
    return text("5 rows\nProduct,Total Sales\nMountain-200,51000\nRoad-150,43000\nTouring-1000,31000\nMountain-100,22000\nRoad-650,18000");
  };
  server.registerTool("run_query", { description: "Runs a DAX query on the dataset and returns the rows as CSV (does not change the user's editor).", inputSchema: z.object({ query: z.string() }), annotations: { readOnlyHint: true } }, async ({ query }) => fakeRun(query));
  server.registerTool("set_editor", { description: "Puts a query in the user's editor (and runs it when run is true).", inputSchema: z.object({ query: z.string(), run: z.boolean().optional() }) }, async ({ query }) => {
    editor = query;
    return text("Query set in the editor.");
  });
  return {
    server,
    check: (outcome) => {
      const problems: string[] = [];
      if (!/TOPN\s*\(\s*5|TOPN\(5/i.test(editor)) problems.push(`editor query not a top 5: ${editor.slice(0, 160)}`);
      if (!/Product'?\[Product\]/.test(editor)) problems.push("editor does not group by 'Product'[Product]");
      if (!queries.length) problems.push("query never tested");
      if (!/Mountain-200/.test(outcome.answer ?? "")) problems.push("answer does not report the result");
      return problems;
    },
  };
}

function researcher(): ReturnType<Scenario["server"]> {
  const server = newServer("pbi-agent", PBI_INSTRUCTIONS);
  let pivot: any = { rows: [], columns: [], values: [], filters: [] };
  page(
    server,
    `Researcher: the user compares data from several sources with a pivot table. Sources (research://sources) are datasets of the database; remote ones speak DAX, local ones SQL.
The pivot (research://pivot) has rows/columns (labels bound to one column per source: "bindings": {sourceId: "'Table'[Column]"}), values (per source: a measure "[Name]" or a column with an aggregate, or a native expression; computed values combine other values by label, e.g. {Sales A} - {Sales B}, DIVIDE({x}, {y}), ABS(…)) and filters (per source, list of values).
To compare two sources, bind each row label to the matching column of both sources and add one value per source plus a computed gap.`,
  );
  server.registerResource("sources", "research://sources", { title: "Sources and their models", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
    contents: [
      {
        uri: uri.href,
        text: `## Source excel-erp: ERP export (SQL)
'Sales'
  columns: Region:String, Month:String, Amount:Double, Orders:Int64
## Source crm-model: CRM (DAX)
'Deals'
  columns: region:String, month:String, revenue:Double
  measures: [Revenue]`,
      },
    ],
  }));
  server.registerResource("pivot", "research://pivot", { title: "Current pivot and its result", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
    contents: [{ uri: uri.href, text: `Pivot:\n${JSON.stringify(pivot)}\n\nResult: ${pivot.values.length ? "4 rows" : "(no result)"}` }],
  }));
  const field = z.string();
  const pivotSchema = z.object({
    rows: z.array(z.object({ label: z.string(), bindings: z.record(z.string(), field) })).default([]),
    columns: z.array(z.object({ label: z.string(), bindings: z.record(z.string(), field) })).default([]),
    values: z
      .array(
        z.object({
          label: z.string(),
          source: z.string().optional(),
          field: field.optional(),
          aggregate: z.enum(["sum", "avg", "count", "distinctcount", "min", "max"]).optional(),
          expression: z.string().optional(),
          computed: z.string().optional(),
          format: z.enum(["number", "integer", "percent", "currency"]).optional(),
        }),
      )
      .default([]),
    filters: z.array(z.object({ source: z.string(), field, values: z.array(z.union([z.string(), z.number(), z.boolean()])) })).default([]),
  });
  server.registerTool("set_pivot", { description: "Replaces the pivot (rows, columns, values, filters). The result appears in research://pivot.", inputSchema: z.object({ pivot: pivotSchema }) }, async ({ pivot: next }) => {
    for (const v of next.values)
      if (v.computed) for (const ref of v.computed.match(/\{([^}]+)\}/g) ?? []) if (!next.values.some((x) => `{${x.label}}` === ref)) return fail(`Unknown value ${ref} (values: ${next.values.map((x) => x.label).join(", ")})`);
    for (const d of [...next.rows, ...next.columns])
      for (const s of next.values.filter((v) => v.source).map((v) => v.source!)) if (!d.bindings[s]) return fail(`"${d.label}" is not mapped to a column of ${s}.`);
    pivot = next;
    return text("4 rows. East: ERP 3630, CRM 3480, gap 150; others equal. See research://pivot.");
  });
  return {
    server,
    check: (outcome) => {
      const problems: string[] = [];
      const region = pivot.rows.find((r: any) => /r[ée]gion/i.test(r.label));
      if (!region || !/Region/.test(region.bindings["excel-erp"] ?? "") || !/region/.test(region.bindings["crm-model"] ?? "")) problems.push(`region not bound to both sources: ${JSON.stringify(pivot.rows)}`);
      if (!pivot.values.some((v: any) => v.source === "excel-erp") || !pivot.values.some((v: any) => v.source === "crm-model")) problems.push("missing a value per source");
      if (!pivot.values.some((v: any) => v.computed)) problems.push("no computed gap");
      if (!/east/i.test(outcome.answer ?? "")) problems.push("answer does not name the region with a gap");
      return problems;
    },
  };
}

/** Large data not inline: the agent must read / search it. */
function largeCatalog(): ReturnType<Scenario["server"]> {
  const server = newServer("pbi-agent", PBI_INSTRUCTIONS);
  page(server, "Database page: lists of reports, datasets (semantic models) and guides stored in local Excel files. Each content_json is readable with the db://{table}/{id}/content template.");
  const reports = Array.from({ length: 1500 }, (_, i) => ({ id: `r${i}`, title: i === 1234 ? "Churn by segment (Q3)" : `Report ${i}`, dataset_id: `d${i % 40}`, context: i === 1234 ? "Owned by the retention team" : "" }));
  server.registerResource("reports", "db://reports", { title: "reports (one row per line)", mimeType: "application/jsonl", annotations: { priority: 1 } }, async (uri) => ({
    contents: [{ uri: uri.href, text: reports.map((r) => JSON.stringify(r)).join("\n") }],
  }));
  server.registerResource("content", new ResourceTemplate("db://{table}/{id}/content", { list: undefined }), { title: "content_json of a report or dataset", mimeType: "application/json" }, async (uri, { id }) => ({
    contents: [{ uri: uri.href, text: JSON.stringify(id === "r1234" ? { pages: [{ displayName: "Overview", visuals: [{ type: "lineChart", title: "Churn rate by month" }, { type: "card", title: "Churned customers" }] }] } : { pages: [] }) }],
  }));
  return {
    server,
    check: (outcome) => {
      const a = outcome.answer ?? "";
      return [!/churn rate by month/i.test(a) && "did not read the report content", !/retention/i.test(a) && "did not report the owner"].filter(Boolean) as string[];
    },
  };
}

export const SCENARIOS: Scenario[] = [
  { name: "excel: schema from a request (FR)", prompt: "fait un schema pour des task avec status", server: schemaEditor },
  { name: "excel: read mode question", prompt: "Combien de tâches sont bloquées ?", server: () => fileEditor("read") },
  { name: "excel: edit rows", prompt: "Mark all of Bob's tasks as done, and add a task 'Write tests' for Alice due next Friday.", server: () => fileEditor("edit") },
  { name: "excel: ambiguous destructive request", prompt: "supprime les tâches", server: () => fileEditor("edit", "ask") },
  {
    name: "excel: follow-up with history",
    history: "USER: Mark all of Bob's tasks as done.\n\nASSISTANT [used: update_rows]: Done: Prepare budget, Call supplier and Fix invoice 42 are now done.",
    prompt: "Oups, remets celle du fournisseur à faire.",
    server: () => {
      const app = fileEditor("edit", "followup");
      return app;
    },
  },
  { name: "pbi: DAX top 5", prompt: "Écris la requête des 5 meilleurs produits par ventes, mets-la dans l'éditeur et dis-moi le résultat.", server: daxRunner },
  { name: "pbi: compare two sources", prompt: "Compare le montant de l'ERP avec le revenue du CRM par région et montre l'écart. Quelle région a un écart ?", server: researcher },
  { name: "pbi: find in large data", prompt: "What is in the report about churn by segment, and who owns it?", server: largeCatalog },
];
