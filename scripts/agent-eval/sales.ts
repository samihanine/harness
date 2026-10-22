/**
 * The live eval's test data: a realistic sales semantic model (star schema: Sales and Returns facts,
 * Date, Product, Customer, Store, Region, monthly Targets; an inactive ship-date relationship; time
 * intelligence measures), created in My workspace with its data embedded in the model ("Enter data"
 * pattern: no source, no credentials), refreshed once, and a report on it for the Builder.
 *
 *   bun run eval:live:setup     creates them (once) → .local/agent-eval/fixtures.json
 *   bun run eval:live:cleanup   deletes everything the setup created
 */
import { deflateRawSync } from "node:zlib";
import { existsSync, mkdirSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { fabric, pbi } from "@/lib/ms";
import { workspaceId } from "@/lib/tmdl";

export const FIXTURES = ".local/agent-eval/fixtures.json";
export type Fixtures = { workspaceId: string; datasetId: string; datasetName: string; reportId?: string; reportName?: string };
export const readFixtures = (): Fixtures | undefined => (existsSync(FIXTURES) ? JSON.parse(readFileSync(FIXTURES, "utf8")) : undefined);

const NAME = "Hub test — Sales (temporaire)";

/* ----------------------------------- Data ----------------------------------- */

/** Deterministic random numbers (same data at every setup). */
let seed = 42;
const rand = () => ((seed = (seed * 1664525 + 1013904223) % 2 ** 32) / 2 ** 32);
const pick = <T,>(list: T[]) => list[Math.floor(rand() * list.length)];
const round = (n: number, d = 2) => Math.round(n * 10 ** d) / 10 ** d;
const iso = (d: Date) => d.toISOString().slice(0, 10);

const REGIONS = [
  { Region: "North America", Manager: "Olivia Martin" },
  { Region: "Europe", Manager: "Lucas Bernard" },
  { Region: "Asia Pacific", Manager: "Mei Chen" },
  { Region: "Latin America", Manager: "Diego Ramos" },
];
const COUNTRIES: Record<string, string[]> = { "North America": ["United States", "Canada"], Europe: ["France", "Germany", "United Kingdom", "Spain"], "Asia Pacific": ["Japan", "Australia", "Singapore"], "Latin America": ["Brazil", "Mexico"] };
const CATALOG: Record<string, Record<string, [number, number]>> = {
  Bikes: { "Road Bikes": [900, 2600], "Mountain Bikes": [700, 2200], "Touring Bikes": [800, 1900] },
  Components: { Frames: [250, 900], Wheels: [80, 400], Brakes: [40, 160] },
  Clothing: { Jerseys: [35, 90], Shorts: [30, 80], Gloves: [15, 45] },
  Accessories: { Helmets: [40, 180], Lights: [15, 70], Bottles: [5, 20] },
  Electronics: { "Bike Computers": [90, 450], "Heart Rate Monitors": [50, 220] },
};
const BRANDS = ["Velocity", "Summit", "Apex", "Trailhead", "Urbana"];

function generate() {
  const products = Object.entries(CATALOG).flatMap(([Category, subs]) =>
    Object.entries(subs).flatMap(([Subcategory, [lo, hi]]) =>
      Array.from({ length: 4 }, (_, i) => {
        const ListPrice = round(lo + rand() * (hi - lo), 0);
        return { ProductKey: 0, Product: `${pick(BRANDS)} ${Subcategory.replace(/s$/, "")} ${["S", "M", "X", "Pro"][i]}`, Category, Subcategory, Brand: "", ListPrice, UnitCost: round(ListPrice * (0.45 + rand() * 0.25)) };
      }),
    ),
  ).map((p, i) => ({ ...p, ProductKey: i + 1, Brand: p.Product.split(" ")[0] }));
  const stores = Array.from({ length: 16 }, (_, i) => {
    const r = REGIONS[i % 4];
    return { StoreKey: i + 1, Store: `${pick(COUNTRIES[r.Region])} ${["Central", "Outlet", "Online", "Wholesale"][Math.floor(i / 4)]}`, Channel: ["Retail", "Retail", "Online", "Wholesale"][Math.floor(i / 4)], Region: r.Region };
  });
  const customers = Array.from({ length: 450 }, (_, i) => {
    const region = pick(REGIONS).Region;
    return { CustomerKey: i + 1, Customer: `Customer ${String(i + 1).padStart(4, "0")}`, Segment: pick(["Consumer", "Consumer", "Corporate", "Small Business"]), Country: pick(COUNTRIES[region]), CustomerRegion: region };
  });
  const sales: Record<string, string | number>[] = [];
  const returns: Record<string, string | number>[] = [];
  const start = new Date("2023-01-01").getTime();
  const days = Math.round((new Date("2026-09-30").getTime() - start) / 864e5);
  for (let o = 1; o <= 3200; o++) {
    // Growth over time and a seasonal peak in spring / before Christmas.
    const t = Math.pow(rand(), 0.8);
    const date = new Date(start + Math.floor(t * days) * 864e5);
    if ([0, 1, 7].includes(date.getMonth()) && rand() < 0.35) continue;
    const customer = pick(customers);
    const store = pick(stores.filter((s) => s.Region === customer.CustomerRegion));
    const lines = 1 + Math.floor(rand() * 3);
    for (let l = 1; l <= lines; l++) {
      const p = pick(products);
      const Quantity = p.Category === "Bikes" ? 1 + Math.floor(rand() * 2) : 1 + Math.floor(rand() * 6);
      const Discount = store.Channel === "Wholesale" ? round(0.1 + rand() * 0.1) : rand() < 0.2 ? 0.1 : 0;
      const SalesAmount = round(Quantity * p.ListPrice * (1 - Discount));
      const ship = new Date(date.getTime() + (1 + Math.floor(rand() * 6)) * 864e5);
      sales.push({ OrderID: `SO${String(o).padStart(5, "0")}`, LineNumber: l, OrderDate: iso(date), ShipDate: iso(ship), ProductKey: p.ProductKey, CustomerKey: customer.CustomerKey, StoreKey: store.StoreKey, Quantity, UnitPrice: p.ListPrice, Discount, SalesAmount, Cost: round(Quantity * p.UnitCost) });
      if (rand() < 0.04) returns.push({ OrderID: `SO${String(o).padStart(5, "0")}`, ProductKey: p.ProductKey, ReturnDate: iso(new Date(ship.getTime() + (3 + Math.floor(rand() * 20)) * 864e5)), Quantity: 1, Reason: pick(["Damaged", "Wrong size", "Not as described", "Changed mind"]) });
    }
  }
  const targets: Record<string, string | number>[] = [];
  for (let y = 2023; y <= 2026; y++)
    for (let m = 0; m < 12; m++) {
      if (y === 2026 && m > 11) break;
      for (const r of REGIONS) {
        const actual = sales.filter((s) => String(s.OrderDate).startsWith(`${y}-${String(m + 1).padStart(2, "0")}`) && stores[Number(s.StoreKey) - 1].Region === r.Region).reduce((a, s) => a + Number(s.SalesAmount), 0);
        targets.push({ MonthStart: `${y}-${String(m + 1).padStart(2, "0")}-01`, Region: r.Region, Target: round((actual || 20000) * (0.95 + rand() * 0.2), -2) });
      }
    }
  return { products, stores, customers, sales, returns, targets, regions: REGIONS };
}

/* ----------------------------------- TMDL ----------------------------------- */

type Col = { name: string; type: "int64" | "double" | "string" | "dateTime"; format?: string; sum?: boolean; key?: boolean; sortBy?: string; hidden?: boolean };
const M_TYPE = { int64: "Int64.Type", double: "type number", string: "type text", dateTime: "type date" };
const q = (n: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

const columnTmdl = (c: Col) =>
  [
    `\tcolumn ${q(c.name)}`,
    `\t\tdataType: ${c.type}`,
    c.key && "\t\tisKey",
    c.hidden && "\t\tisHidden",
    c.format && `\t\tformatString: ${c.format}`,
    `\t\tsummarizeBy: ${c.sum ? "sum" : "none"}`,
    `\t\tsourceColumn: ${c.name}`,
    c.sortBy && `\t\tsortByColumn: ${q(c.sortBy)}`,
  ]
    .filter(Boolean)
    .join("\n");

/** A table whose rows are embedded in the model (compressed JSON, as Power BI's "Enter data"). */
function dataTable(name: string, cols: Col[], rows: Record<string, unknown>[], measures = "") {
  const json = JSON.stringify(rows.map((r) => cols.map((c) => (r[c.name] === undefined || r[c.name] === null ? null : String(r[c.name])))));
  const base64 = deflateRawSync(Buffer.from(json)).toString("base64");
  return `table ${q(name)}

${cols.map(columnTmdl).join("\n\n")}
${measures}
\tpartition ${q(name)} = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet
\t\t\t\t    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("${base64}", BinaryEncoding.Base64), Compression.Deflate)), {${cols.map((c) => `"${c.name}"`).join(", ")}}),
\t\t\t\t    Typed = Table.TransformColumnTypes(Source, {${cols.map((c) => `{"${c.name}", ${M_TYPE[c.type]}}`).join(", ")}}, "en-US")
\t\t\t\tin
\t\t\t\t    Typed
`;
}

const measure = (name: string, dax: string, format: string, folder?: string, description?: string) =>
  [description && `\t/// ${description}`, `\tmeasure ${q(name)} = ${dax}`, `\t\tformatString: ${format}`, folder && `\t\tdisplayFolder: ${folder}`].filter(Boolean).join("\n");

const MEASURES = [
  measure("Total Sales", "SUM(Sales[SalesAmount])", "\\$#,0", "Sales", "Invoiced amount after discount"),
  measure("Total Cost", "SUM(Sales[Cost])", "\\$#,0", "Sales"),
  measure("Gross Margin", "[Total Sales] - [Total Cost]", "\\$#,0", "Margin"),
  measure("Margin %", "DIVIDE([Gross Margin], [Total Sales])", "0.0%", "Margin"),
  measure("Orders", "DISTINCTCOUNT(Sales[OrderID])", "#,0", "Sales"),
  measure("Units", "SUM(Sales[Quantity])", "#,0", "Sales"),
  measure("Avg Order Value", "DIVIDE([Total Sales], [Orders])", "\\$#,0.00", "Sales"),
  measure("Active Customers", "DISTINCTCOUNT(Sales[CustomerKey])", "#,0", "Customers"),
  measure("Avg Discount", "AVERAGE(Sales[Discount])", "0.0%", "Sales"),
  measure("Sales LY", "CALCULATE([Total Sales], SAMEPERIODLASTYEAR('Date'[Date]))", "\\$#,0", "Time intelligence", "Sales of the same period last year"),
  measure("Sales YoY %", "DIVIDE([Total Sales] - [Sales LY], [Sales LY])", "0.0%", "Time intelligence"),
  measure("Sales YTD", "TOTALYTD([Total Sales], 'Date'[Date])", "\\$#,0", "Time intelligence"),
  measure("Shipped Sales", "CALCULATE([Total Sales], USERELATIONSHIP(Sales[ShipDate], 'Date'[Date]))", "\\$#,0", "Time intelligence", "Sales by ship date instead of order date"),
  measure("Target", "SUM(Targets[Target])", "\\$#,0", "Targets"),
  measure("Sales vs Target %", "DIVIDE([Total Sales], [Target])", "0.0%", "Targets"),
  measure("Returned Units", "SUM(Returns[Quantity])", "#,0", "Returns"),
  measure("Return Rate", "DIVIDE([Returned Units], [Units])", "0.00%", "Returns"),
].join("\n\n");

const DATE_TABLE = `table Date
\tdataCategory: Time

\tcolumn Date
\t\tdataType: dateTime
\t\tisKey
\t\tformatString: yyyy-mm-dd
\t\tsummarizeBy: none
\t\tsourceColumn: Date

\tcolumn Year
\t\tdataType: int64
\t\tformatString: 0
\t\tsummarizeBy: none
\t\tsourceColumn: Year

\tcolumn Quarter
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Quarter

\tcolumn 'Month Number'
\t\tdataType: int64
\t\tisHidden
\t\tsummarizeBy: none
\t\tsourceColumn: Month Number

\tcolumn Month
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Month
\t\tsortByColumn: 'Month Number'

\tcolumn 'Year Month'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Year Month

\tcolumn 'Month Start'
\t\tdataType: dateTime
\t\tformatString: yyyy-mm-dd
\t\tsummarizeBy: none
\t\tsourceColumn: Month Start

\tcolumn Weekday
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Weekday

\tpartition Date = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet
\t\t\t\t    Dates = List.Dates(#date(2023, 1, 1), Duration.Days(#date(2026, 12, 31) - #date(2023, 1, 1)) + 1, #duration(1, 0, 0, 0)),
\t\t\t\t    Base = Table.TransformColumnTypes(Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}), {{"Date", type date}}),
\t\t\t\t    Year = Table.AddColumn(Base, "Year", each Date.Year([Date]), Int64.Type),
\t\t\t\t    Quarter = Table.AddColumn(Year, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
\t\t\t\t    MonthNumber = Table.AddColumn(Quarter, "Month Number", each Date.Month([Date]), Int64.Type),
\t\t\t\t    Month = Table.AddColumn(MonthNumber, "Month", each Date.ToText([Date], "MMM", "en-US"), type text),
\t\t\t\t    YearMonth = Table.AddColumn(Month, "Year Month", each Date.ToText([Date], "yyyy-MM"), type text),
\t\t\t\t    MonthStart = Table.AddColumn(YearMonth, "Month Start", each Date.StartOfMonth([Date]), type date),
\t\t\t\t    Weekday = Table.AddColumn(MonthStart, "Weekday", each Date.ToText([Date], "ddd", "en-US"), type text)
\t\t\t\tin
\t\t\t\t    Weekday
`;

const RELATIONSHIPS = [
  ["Sales_OrderDate", "Sales.OrderDate", "Date.Date"],
  ["Sales_ShipDate", "Sales.ShipDate", "Date.Date", "inactive"],
  ["Sales_Product", "Sales.ProductKey", "Product.ProductKey"],
  ["Sales_Customer", "Sales.CustomerKey", "Customer.CustomerKey"],
  ["Sales_Store", "Sales.StoreKey", "Store.StoreKey"],
  ["Store_Region", "Store.Region", "Region.Region"],
  ["Targets_Month", "Targets.MonthStart", "Date.Date"],
  ["Targets_Region", "Targets.Region", "Region.Region"],
  ["Returns_Product", "Returns.ProductKey", "Product.ProductKey"],
  ["Returns_Date", "Returns.ReturnDate", "Date.Date"],
]
  .map(([name, from, to, inactive]) => `relationship ${name}\n${inactive ? "\tisActive: false\n" : ""}\tfromColumn: ${from}\n\ttoColumn: ${to}\n`)
  .join("\n");

function modelParts() {
  const d = generate();
  const tables: Record<string, string> = {
    Sales: dataTable(
      "Sales",
      [
        { name: "OrderID", type: "string" },
        { name: "LineNumber", type: "int64" },
        { name: "OrderDate", type: "dateTime", format: "yyyy-mm-dd" },
        { name: "ShipDate", type: "dateTime", format: "yyyy-mm-dd" },
        { name: "ProductKey", type: "int64", hidden: true },
        { name: "CustomerKey", type: "int64", hidden: true },
        { name: "StoreKey", type: "int64", hidden: true },
        { name: "Quantity", type: "int64", sum: true },
        { name: "UnitPrice", type: "double", format: "\\$#,0.00" },
        { name: "Discount", type: "double", format: "0%" },
        { name: "SalesAmount", type: "double", format: "\\$#,0.00", sum: true },
        { name: "Cost", type: "double", format: "\\$#,0.00", sum: true },
      ],
      d.sales,
      `\n${MEASURES}\n`,
    ),
    Returns: dataTable("Returns", [{ name: "OrderID", type: "string" }, { name: "ProductKey", type: "int64", hidden: true }, { name: "ReturnDate", type: "dateTime", format: "yyyy-mm-dd" }, { name: "Quantity", type: "int64", sum: true }, { name: "Reason", type: "string" }], d.returns),
    Product: dataTable("Product", [{ name: "ProductKey", type: "int64", key: true, hidden: true }, { name: "Product", type: "string" }, { name: "Category", type: "string" }, { name: "Subcategory", type: "string" }, { name: "Brand", type: "string" }, { name: "ListPrice", type: "double", format: "\\$#,0.00" }, { name: "UnitCost", type: "double", format: "\\$#,0.00" }], d.products),
    Customer: dataTable("Customer", [{ name: "CustomerKey", type: "int64", key: true, hidden: true }, { name: "Customer", type: "string" }, { name: "Segment", type: "string" }, { name: "Country", type: "string" }, { name: "CustomerRegion", type: "string" }], d.customers),
    Store: dataTable("Store", [{ name: "StoreKey", type: "int64", key: true, hidden: true }, { name: "Store", type: "string" }, { name: "Channel", type: "string" }, { name: "Region", type: "string" }], d.stores),
    Region: dataTable("Region", [{ name: "Region", type: "string", key: true }, { name: "Manager", type: "string" }], d.regions),
    Targets: dataTable("Targets", [{ name: "MonthStart", type: "dateTime", format: "yyyy-mm-dd" }, { name: "Region", type: "string" }, { name: "Target", type: "double", format: "\\$#,0", sum: true }], d.targets),
    Date: DATE_TABLE,
  };
  const files = [
    { path: "definition.pbism", text: JSON.stringify({ version: "4.2", settings: {} }) },
    { path: "definition/database.tmdl", text: "database\n\tcompatibilityLevel: 1567\n" },
    { path: "definition/model.tmdl", text: `model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: en-US\n\n${Object.keys(tables).map((t) => `ref table ${q(t)}`).join("\n")}\n` },
    { path: "definition/relationships.tmdl", text: RELATIONSHIPS },
    ...Object.entries(tables).map(([name, text]) => ({ path: `definition/tables/${name}.tmdl`, text })),
  ];
  console.log(`Data: ${d.sales.length} sales lines, ${d.returns.length} returns, ${d.products.length} products, ${d.customers.length} customers, ${d.targets.length} targets.`);
  return files.map((f) => ({ path: f.path, payload: Buffer.from(f.text).toString("base64"), payloadType: "InlineBase64" }));
}

/* ------------------------------- Setup / cleanup ------------------------------ */

async function setup() {
  if (readFixtures()) return console.log(`Already set up (${FIXTURES}): run eval:live:cleanup first to start again.`);
  const ws = await workspaceId();
  console.log("Creating the semantic model…");
  const created = await fabric(`/workspaces/${ws}/semanticModels`, "POST", { displayName: NAME, definition: { parts: modelParts() } });
  const datasetId = created?.id ?? (await fabric(`/workspaces/${ws}/semanticModels`)).value.find((m: { displayName: string }) => m.displayName === NAME)?.id;
  const fixtures: Fixtures = { workspaceId: ws, datasetId, datasetName: NAME };
  mkdirSync(".local/agent-eval", { recursive: true });
  writeFileSync(FIXTURES, JSON.stringify(fixtures, null, 2));
  console.log(`Model ${datasetId}. Refreshing it once (the data is in the model)…`);
  const { refreshModel } = await import("@/server/ms");
  // My workspace: the Power BI API has no /groups path for it.
  const result = await refreshModel({ data: { datasetId } });
  console.log(`Refresh: ${result.status}${result.error ? ` — ${result.error}` : ""}`);
  if (result.status !== "Completed") throw new Error("The refresh did not complete: see the error above.");
  console.log("Done:", FIXTURES);
}

async function cleanup() {
  const f = readFixtures();
  if (!f) return console.log("Nothing to clean up.");
  if (f.reportId) await fabric(`/workspaces/${f.workspaceId}/reports/${f.reportId}`, "DELETE").catch((e) => console.log("report:", e.message));
  // Reports made by the scenarios on the test model (research reports).
  const reports = (await pbi(`/reports`).catch(() => ({ value: [] }))).value as { id: string; name: string; datasetId: string }[];
  for (const r of reports.filter((x) => x.datasetId === f.datasetId)) await fabric(`/workspaces/${f.workspaceId}/reports/${r.id}`, "DELETE").catch((e) => console.log(r.name, e.message));
  await fabric(`/workspaces/${f.workspaceId}/semanticModels/${f.datasetId}`, "DELETE");
  unlinkSync(FIXTURES);
  console.log("Test model and its reports deleted.");
}

if (import.meta.main) {
  const command = process.argv[2];
  if (command === "setup") await setup();
  else if (command === "cleanup") await cleanup();
  else console.log("bun run eval:live:setup | eval:live:cleanup");
}
