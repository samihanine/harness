/**
 * The Researcher's working workbook, written by this code (never typed by the AI): one sheet per
 * analysis — a PivotTable connected live to a Power BI semantic model (OLAP, computed by Excel when the
 * file opens), or an Excel table whose columns are formulas. Only the definitions are written: Excel
 * computes the values.
 */
import JSZip from "jszip";

/** A field of a pivot: a column of the model ('Table'[Column]) on rows / columns / filters, or a measure in values. */
export type PivotField = { table: string; column: string };
export type PivotSheet = {
  kind: "pivot";
  name: string;
  dataset: { id: string; name: string };
  rows: PivotField[];
  columns: PivotField[];
  values: { table: string; measure: string }[];
  /** Filters: one value kept per field (as it is in the data). */
  filters: (PivotField & { value?: string })[];
};
export type FormulaSheet = {
  kind: "table";
  name: string;
  /** Number of rows of the table. */
  rows: number;
  columns: { name: string; formula?: string }[];
};
/** A chart of another sheet's table or pivot (added by Excel on its own sheet once the file is published). */
export type ChartSheet = { kind: "chart"; name: string; source: string; type: ChartType };
export const CHART_TYPES = ["ColumnClustered", "ColumnStacked", "BarClustered", "BarStacked", "Line", "Area", "Pie", "Doughnut"] as const;
export type ChartType = (typeof CHART_TYPES)[number];
export type Sheet = PivotSheet | FormulaSheet;
export type ResearchSheet = Sheet | ChartSheet;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const mdx = (s: string) => `[${s.replace(/]/g, "]]")}]`;
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const xml = (body: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;

/** Sheet name as Excel allows it (31 characters, no []:*?/\). */
export const sheetName = (s: string) => s.replace(/[[\]:*?/\\]/g, " ").trim().slice(0, 31) || "Sheet";
const letter = (n: number) => {
  let s = "";
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

/** The connection of a pivot to a semantic model (the same as "Analyze in Excel"): the user's own identity, no secret in the file. */
function connection(id: number, dataset: { id: string; name: string }, tenant: string) {
  const conn = [
    "Provider=MSOLAP.8",
    "Integrated Security=ClaimsToken",
    "Persist Security Info=True",
    `Initial Catalog=sobe_wowvirtualserver-${dataset.id}`,
    "Data Source=pbiazure://api.powerbi.com",
    "MDX Compatibility=1",
    "Safety Options=2",
    "MDX Missing Member Mode=Error",
    `Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, ${tenant}`,
    "Update Isolation Level=2",
  ].join(";");
  return `<connection id="${id}" keepAlive="1" name="${esc(`Power BI ${dataset.name}`)}" type="5" refreshedVersion="8" background="1" refreshOnLoad="1"><dbPr connection="${esc(conn)}" command="Model" commandType="1"/><olapPr sendLocale="1" rowDrillCount="1000"/></connection>`;
}

/**
 * Cache and table of an OLAP pivot: the hierarchies used ([Table].[Column] attributes, [Measures].[M]),
 * no data (refreshed when the file opens).
 */
function pivotParts(s: PivotSheet, cacheId: number, connectionId: number) {
  const attrs: (PivotField & { axis?: "axisRow" | "axisCol" | "axisPage"; value?: string })[] = [];
  const add = (f: PivotField & { value?: string }, axis: "axisRow" | "axisCol" | "axisPage") => {
    if (!attrs.some((a) => a.table === f.table && a.column === f.column)) attrs.push({ ...f, axis });
  };
  s.rows.forEach((f) => add(f, "axisRow"));
  s.columns.forEach((f) => add(f, "axisCol"));
  s.filters.forEach((f) => add(f, "axisPage"));
  const tables = [...new Set([...attrs.map((a) => a.table), ...s.values.map((v) => v.table)])];
  // Hierarchies: attributes first, then measures; cache fields in the same order.
  const hierarchies = [
    ...attrs.map((a, i) => `<cacheHierarchy uniqueName="${esc(`${mdx(a.table)}.${mdx(a.column)}`)}" caption="${esc(a.column)}" attribute="1" defaultMemberUniqueName="${esc(`${mdx(a.table)}.${mdx(a.column)}.[All]`)}" allUniqueName="${esc(`${mdx(a.table)}.${mdx(a.column)}.[All]`)}" dimensionUniqueName="${esc(mdx(a.table))}" displayFolder="" count="2" unbalanced="0"><fieldsUsage count="2"><fieldUsage x="-1"/><fieldUsage x="${i}"/></fieldsUsage></cacheHierarchy>`),
    ...s.values.map((v, i) => `<cacheHierarchy uniqueName="${esc(`[Measures].${mdx(v.measure)}`)}" caption="${esc(v.measure)}" measure="1" displayFolder="" measureGroup="${esc(v.table)}" count="0" oneField="1"><fieldsUsage count="1"><fieldUsage x="${attrs.length + i}"/></fieldsUsage></cacheHierarchy>`),
  ];
  const fields = [
    ...attrs.map((a, i) => `<cacheField name="${esc(`${mdx(a.table)}.${mdx(a.column)}.${mdx(a.column)}`)}" caption="${esc(a.column)}" numFmtId="0" hierarchy="${i}" level="1"><sharedItems count="0"/></cacheField>`),
    ...s.values.map((v, i) => `<cacheField name="${esc(`[Measures].${mdx(v.measure)}`)}" caption="${esc(v.measure)}" numFmtId="0" hierarchy="${attrs.length + i}" level="32767"/>`),
  ];
  const cache = xml(
    `<pivotCacheDefinition ${NS} saveData="0" refreshOnLoad="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="0" supportSubquery="1" supportAdvancedDrill="1">` +
      `<cacheSource type="external" connectionId="${connectionId}"/>` +
      `<cacheFields count="${fields.length}">${fields.join("")}</cacheFields>` +
      `<cacheHierarchies count="${hierarchies.length}">${hierarchies.join("")}</cacheHierarchies>` +
      `<kpis count="0"/>` +
      `<dimensions count="${tables.length + 1}"><dimension measure="1" name="Measures" uniqueName="[Measures]" caption="Measures"/>${tables.map((t) => `<dimension name="${esc(t)}" uniqueName="${esc(mdx(t))}" caption="${esc(t)}"/>`).join("")}</dimensions>` +
      `<measureGroups count="${tables.length}">${tables.map((t) => `<measureGroup name="${esc(t)}" caption="${esc(t)}"/>`).join("")}</measureGroups>` +
      `<maps count="${tables.length}">${tables.map((_, i) => `<map measureGroup="${i}" dimension="${i + 1}"/>`).join("")}</maps>` +
      `</pivotCacheDefinition>`,
  );
  const rowsIdx = attrs.map((a, i) => (a.axis === "axisRow" ? i : -1)).filter((i) => i >= 0);
  const colsIdx = attrs.map((a, i) => (a.axis === "axisCol" ? i : -1)).filter((i) => i >= 0);
  const pageIdx = attrs.map((a, i) => (a.axis === "axisPage" ? i : -1)).filter((i) => i >= 0);
  const pivotFields = [
    ...attrs.map((a) => `<pivotField axis="${a.axis}" allDrilled="1" showAll="0" dataSourceSort="1" defaultAttributeDrillState="1"><items count="1"><item t="default"/></items></pivotField>`),
    ...s.values.map(() => `<pivotField dataField="1" showAll="0"/>`),
  ];
  const top = pageIdx.length ? pageIdx.length + 2 : 1;
  const table = xml(
    `<pivotTableDefinition ${NS} name="${esc(sheetName(s.name))}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" fieldListSortAscending="1">` +
      `<location ref="A${top}:${letter(Math.max(1, colsIdx.length ? 2 : s.values.length))}${top + 1}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"${pageIdx.length ? ` rowPageCount="${pageIdx.length}" colPageCount="1"` : ""}/>` +
      `<pivotFields count="${pivotFields.length}">${pivotFields.join("")}</pivotFields>` +
      (rowsIdx.length ? `<rowFields count="${rowsIdx.length}">${rowsIdx.map((i) => `<field x="${i}"/>`).join("")}</rowFields><rowItems count="1"><i t="grand"><x/></i></rowItems>` : "") +
      (colsIdx.length || s.values.length > 1 ? `<colFields count="${colsIdx.length + (s.values.length > 1 ? 1 : 0)}">${colsIdx.map((i) => `<field x="${i}"/>`).join("")}${s.values.length > 1 ? `<field x="-2"/>` : ""}</colFields>` : "") +
      `<colItems count="1"><i/></colItems>` +
      (pageIdx.length
        ? `<pageFields count="${pageIdx.length}">${pageIdx.map((i) => {
            const a = attrs[i];
            return `<pageField fld="${i}" hier="${i}"${a.value !== undefined ? ` name="${esc(`${mdx(a.table)}.${mdx(a.column)}.&${mdx(a.value)}`)}" cap="${esc(a.value)}"` : ""}/>`;
          }).join("")}</pageFields>`
        : "") +
      (s.values.length ? `<dataFields count="${s.values.length}">${s.values.map((_, i) => `<dataField fld="${attrs.length + i}" baseField="0" baseItem="0"/>`).join("")}</dataFields>` : "") +
      // One per cache hierarchy (OLAP pivots need it): measures can go to the values.
      `<pivotHierarchies count="${attrs.length + s.values.length}">${attrs.map(() => "<pivotHierarchy/>").join("")}${s.values.map(() => '<pivotHierarchy dragToData="1"/>').join("")}</pivotHierarchies>` +
      `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>` +
      `<rowHierarchiesUsage count="${rowsIdx.length}">${rowsIdx.map((i) => `<rowHierarchyUsage hierarchyUsage="${i}"/>`).join("")}</rowHierarchiesUsage>` +
      (colsIdx.length ? `<colHierarchiesUsage count="${colsIdx.length}">${colsIdx.map((i) => `<colHierarchyUsage hierarchyUsage="${i}"/>`).join("")}</colHierarchiesUsage>` : "") +
      `</pivotTableDefinition>`,
  );
  return { cache, table };
}

/**
 * A formula as the file stores it: the "@" shorthand of Excel's interface ([@Col], Table[@Col],
 * [@[Col name]]) is written [[#This Row],[Col]] with its table name.
 */
export function fileFormula(formula: string, table: string) {
  return formula
    .replace(/^=/, "")
    .replace(/(\b[\p{L}_][\p{L}\p{N}_.]*)?\[@(\[[^\]]+\]|[^\]]+)\]/gu, (_, t: string | undefined, col: string) => `${t ?? table}[[#This Row],${col.startsWith("[") ? col : `[${col}]`}]`);
}

/** A table whose columns are formulas (structured references to the other tables of the workbook). */
function formulaParts(s: FormulaSheet, id: number) {
  const rows = Math.max(1, s.rows);
  const ref = `A1:${letter(s.columns.length - 1)}${rows + 1}`;
  const header = `<row r="1">${s.columns.map((c, i) => `<c r="${letter(i)}1" t="inlineStr"><is><t>${esc(c.name)}</t></is></c>`).join("")}</row>`;
  const body = Array.from({ length: rows }, (_, r) => `<row r="${r + 2}">${s.columns.map((c, i) => (c.formula ? `<c r="${letter(i)}${r + 2}"><f>${esc(fileFormula(c.formula, tableName(s.name)))}</f></c>` : "")).join("")}</row>`).join("");
  const table = xml(
    `<table ${NS} id="${id}" name="${esc(tableName(s.name))}" displayName="${esc(tableName(s.name))}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/>` +
      `<tableColumns count="${s.columns.length}">${s.columns.map((c, i) => `<tableColumn id="${i + 1}" name="${esc(c.name)}">${c.formula ? `<calculatedColumnFormula>${esc(fileFormula(c.formula, tableName(s.name)))}</calculatedColumnFormula>` : ""}</tableColumn>`).join("")}</tableColumns>` +
      `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`,
  );
  return { sheetData: header + body, table, ref };
}
/** Excel table name of a sheet (letters, digits, _; starts with a letter). */
export const tableName = (s: string) => {
  const n = s.replace(/[^\p{L}\p{N}_]+/gu, "_").replace(/^_+|_+$/g, "");
  return /^\p{L}/u.test(n) ? n.slice(0, 60) : `T_${n}`.slice(0, 60);
};

/** The whole workbook (.xlsx) of the sheets. `tenant`: the organization's id (for the connections). */
export async function buildWorkbook(sheets: ResearchSheet[], tenant: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const list = sheets.length ? sheets : [{ kind: "chart", name: "Start", source: "", type: "Line" } as ChartSheet];
  const ct: string[] = [];
  const wbRels: string[] = [];
  const pivots: string[] = [];
  const connections: string[] = [];
  const datasets = new Map<string, number>();
  list.forEach((s, i) => {
    const n = i + 1;
    const rels: string[] = [];
    let sheetData = "";
    if (s.kind === "pivot") {
      if (!datasets.has(s.dataset.id)) {
        datasets.set(s.dataset.id, datasets.size + 1);
        connections.push(connection(datasets.get(s.dataset.id)!, s.dataset, tenant));
      }
      const { cache, table } = pivotParts(s, n, datasets.get(s.dataset.id)!);
      zip.file(`xl/pivotCache/pivotCacheDefinition${n}.xml`, cache);
      zip.file(`xl/pivotTables/pivotTable${n}.xml`, table);
      zip.file(`xl/pivotTables/_rels/pivotTable${n}.xml.rels`, xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`));
      ct.push(`<Override PartName="/xl/pivotCache/pivotCacheDefinition${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`, `<Override PartName="/xl/pivotTables/pivotTable${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`);
      wbRels.push(`<Relationship Id="rIdCache${n}" Type="${REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${n}.xml"/>`);
      pivots.push(`<pivotCache cacheId="${n}" r:id="rIdCache${n}"/>`);
      rels.push(`<Relationship Id="rId1" Type="${REL}/pivotTable" Target="../pivotTables/pivotTable${n}.xml"/>`);
    } else if (s.kind === "table") {
      const { sheetData: data, table } = formulaParts(s, n);
      sheetData = data;
      zip.file(`xl/tables/table${n}.xml`, table);
      ct.push(`<Override PartName="/xl/tables/table${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
      rels.push(`<Relationship Id="rId1" Type="${REL}/table" Target="../tables/table${n}.xml"/>`);
    }
    // A chart sheet stays empty here: its chart is added by Excel once the file is published.
    zip.file(
      `xl/worksheets/sheet${n}.xml`,
      xml(`<worksheet ${NS}><sheetViews><sheetView workbookViewId="0"${n === 1 ? ' tabSelected="1"' : ""}/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="12" width="22" customWidth="1"/></cols><sheetData>${sheetData}</sheetData>${s.kind === "table" ? `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>` : ""}</worksheet>`),
    );
    zip.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`));
    ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    wbRels.push(`<Relationship Id="rIdSheet${n}" Type="${REL}/worksheet" Target="worksheets/sheet${n}.xml"/>`);
  });
  zip.file(
    "xl/workbook.xml",
    xml(
      `<workbook ${NS}><bookViews><workbookView activeTab="0"/></bookViews><sheets>${list.map((s, i) => `<sheet name="${esc(sheetName(s.name))}" sheetId="${i + 1}" r:id="rIdSheet${i + 1}"/>`).join("")}</sheets>${pivots.length ? `<pivotCaches>${pivots.join("")}</pivotCaches>` : ""}</workbook>`,
    ),
  );
  zip.file("xl/styles.xml", xml(`<styleSheet ${NS}><fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>`));
  wbRels.push(`<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>`);
  if (connections.length) {
    zip.file("xl/connections.xml", xml(`<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${connections.join("")}</connections>`));
    wbRels.push(`<Relationship Id="rIdConnections" Type="${REL}/connections" Target="connections.xml"/>`);
    ct.push(`<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>`);
  }
  zip.file("xl/_rels/workbook.xml.rels", xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join("")}</Relationships>`));
  zip.file("_rels/.rels", xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`));
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ct.join("")}</Types>`,
    ),
  );
  return zip.generateAsync({ type: "uint8array" });
}
