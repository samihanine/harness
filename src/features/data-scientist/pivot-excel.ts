/**
 * A research's workbook as an Excel file, each pivot connected to its semantic model. A pivot of model
 * measures by columns, filtered on values, is a PivotTable ON the model (an OLAP PivotTable, as "Analyze
 * in Excel" makes: the model's field list, refreshed from Power BI). Any other (aggregated columns,
 * ranges, exclusions): a "… data" sheet holding its query as a table linked to the model (an OLE DB
 * connection running the sheet's DAX) and the sheet itself with a PivotTable on that table. Calculation sheets keep their
 * formulas (GETPIVOTDATA pointed at the pivots). Written by hand (the libraries in reach do not write
 * pivot tables); the rows already read and the computed values fill the file, so it opens with the figures.
 */
import JSZip from "jszip";
import type { Rows } from "@/services/pbi";
import type { ModelInfo } from "@/services/model";
import type { Result } from "./formulas";
import { activeFilters, pivot, resultName, type CalcSheet, type PivotSpec } from "./research";

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const esc = (v: unknown) => String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
/** Column letters of a 0-based index (0 → A). */
const col = (i: number): string => (i >= 26 ? col(Math.floor(i / 26) - 1) : "") + String.fromCharCode(65 + (i % 26));
const isNumber = (v: unknown) => typeof v === "number" && Number.isFinite(v);

/** Power BI's connection for Excel ("Analyze in Excel": works on Pro), its tenant from the token. */
export function powerBiConnection(datasetId: string, tenantId?: string) {
  return [
    "Provider=MSOLAP.8",
    "Integrated Security=ClaimsToken",
    "Persist Security Info=True",
    `Initial Catalog=sobe_wowvirtualserver-${datasetId}`,
    "Data Source=pbiazure://api.powerbi.com",
    "MDX Compatibility=1",
    "Safety Options=2",
    "MDX Missing Member Mode=Error",
    `Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, ${tenantId ?? "common"}`,
    "Update Isolation Level=2",
  ].join(";");
}

/** The tenant of a Microsoft token (its `tid` claim). */
export function tenantOf(token: string) {
  try {
    return JSON.parse(atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))).tid as string | undefined;
  } catch {
    return undefined;
  }
}


const cell = (v: unknown, ref: string) => (isNumber(v) ? `<c r="${ref}"><v>${v}</v></c>` : v === null || v === undefined || v === "" ? "" : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`);

/** A pivot of the workbook: its sheet name, parameters, rows, DAX and model connection. */
export type PivotPart = { name: string; spec: PivotSpec; rows: Rows; dax: string; connection: string; model: string; info?: ModelInfo; note?: string };

/* ------------------------------- On the model (OLAP) ------------------------------- */

const bracket = (s: string) => `[${s.replace(/\]/g, "]]")}]`;
/** 'T'[col] (or T[col]) → its table and column. */
const columnOf = (field: string) => field.trim().match(/^'?(.+?)'?\[([^\]]+)\]$/)?.slice(1, 3) as [string, string] | undefined;
/** [Measure] → its name. */
const measureOf = (value: string) => value.trim().match(/^\[([^\]]+)\]$/)?.[1];

/** A pivot as the model names it: hierarchies of its fields, measures of its values (undefined: not reachable as OLAP). */
export type Olap = {
  /** By the field's name in the rows (resultName): the hierarchy, e.g. [Zone].[ZoneCode]. */
  hierarchies: Map<string, string>;
  /** By the value's name: the measure's unique name, e.g. [Measures].[Sales]. */
  measures: Map<string, string>;
};

/**
 * The model's names of a pivot when Excel can build it on the model: its values are measures of the model,
 * its rows and columns columns of the model, its filters lists of values (in).
 */
export function olapOf(given: PivotSpec, info: ModelInfo | undefined): (Olap & { groups: Map<string, string> }) | undefined {
  if (!info) return undefined;
  const spec = { ...given, filters: activeFilters(given.filters) };
  const hierarchies = new Map<string, string>();
  for (const f of [...spec.rows, ...spec.columns, ...spec.filters.map((x) => x.field)]) {
    const c = columnOf(f);
    const table = c && info.tables.find((t) => t.name.toLowerCase() === c[0].toLowerCase());
    const column = table?.columns.find((x) => x.name.toLowerCase() === c![1].toLowerCase());
    if (!table || !column) return undefined;
    hierarchies.set(column.name, `${bracket(table.name)}.${bracket(column.name)}`);
  }
  if (spec.filters.some((f) => (f.op ?? "in") !== "in" || !f.values?.length)) return undefined;
  const measures = new Map<string, string>();
  const groups = new Map<string, string>();
  for (const v of spec.values) {
    const name = measureOf(v);
    const table = name && info.tables.find((t) => t.measures.some((m) => m.name.toLowerCase() === name.toLowerCase()));
    if (!name || !table) return undefined;
    const real = table.measures.find((m) => m.name.toLowerCase() === name.toLowerCase())!.name;
    measures.set(real, `[Measures].${bracket(real)}`);
    groups.set(real, table.name);
  }
  return measures.size ? { hierarchies, measures, groups } : undefined;
}

/** A member of a hierarchy by its value, e.g. [Zone].[ZoneCode].&[CHI]. */
const member = (hierarchy: string, v: unknown) => `${hierarchy}.&${bracket(String(v))}`;

/**
 * The parts of a pivot on the model (n from 1): its connection to the model (the cube "Model"), its cache
 * (the hierarchies and measures it uses, the members already read) and the PivotTable, refreshed when the
 * file opens. The figures already read fill its cells, so the sheet shows them before the refresh.
 */
function olapParts(n: number, { name, spec: given, rows, connection, model, note }: PivotPart, olap: NonNullable<ReturnType<typeof olapOf>>) {
  const spec = { ...given, filters: activeFilters(given.filters) };
  const rowKeys = spec.rows.map(resultName);
  const colKeys = spec.columns.map(resultName);
  const pageKeys = spec.filters.map((f) => resultName(f.field)).filter((k) => !rowKeys.includes(k) && !colKeys.includes(k));
  const fieldKeys = [...new Set([...rowKeys, ...colKeys, ...pageKeys])];
  const measureKeys = [...olap.measures.keys()];
  // Cache fields: one per field (its level) then one per measure; cache hierarchies in the same order.
  const fieldNo = (k: string) => fieldKeys.indexOf(k);
  const measureNo = (m: string) => fieldKeys.length + measureKeys.indexOf(m);
  const membersOf = (k: string) => {
    const filter = spec.filters.find((f) => resultName(f.field) === k);
    const values = rowKeys.includes(k) || colKeys.includes(k) ? rows.map((r) => r[k]) : (filter?.values ?? []);
    return [...new Set(values.filter((v) => v !== null && v !== undefined && v !== "").map(String))];
  };
  const level = (k: string) => `${olap.hierarchies.get(k)}.${bracket(k)}`;
  const cacheFields = [
    ...fieldKeys.map((k) => {
      const list = membersOf(k);
      return `<cacheField name="${esc(level(k))}" caption="${esc(k)}" numFmtId="0" hierarchy="${fieldNo(k)}" level="1"><sharedItems count="${list.length}">${list.map((v) => `<s v="${esc(member(olap.hierarchies.get(k)!, v))}" c="${esc(v)}"/>`).join("")}</sharedItems></cacheField>`;
    }),
    ...measureKeys.map((m) => `<cacheField name="${esc(olap.measures.get(m))}" caption="${esc(m)}" numFmtId="0" hierarchy="${measureNo(m)}" level="32767"/>`),
  ];
  const cacheHierarchies = [
    ...fieldKeys.map((k) => {
      const h = olap.hierarchies.get(k)!;
      const dimension = h.slice(0, h.indexOf("].[") + 1);
      return `<cacheHierarchy uniqueName="${esc(h)}" caption="${esc(k)}" attribute="1" defaultMemberUniqueName="${esc(`${h}.[All]`)}" allUniqueName="${esc(`${h}.[All]`)}" dimensionUniqueName="${esc(dimension)}" displayFolder="" count="2" unbalanced="0"><fieldsUsage count="2"><fieldUsage x="-1"/><fieldUsage x="${fieldNo(k)}"/></fieldsUsage></cacheHierarchy>`;
    }),
    ...measureKeys.map((m) => `<cacheHierarchy uniqueName="${esc(olap.measures.get(m))}" caption="${esc(m)}" measure="1" displayFolder="" measureGroup="${esc(olap.groups.get(m))}" count="0" oneField="1"><fieldsUsage count="1"><fieldUsage x="${measureNo(m)}"/></fieldsUsage></cacheHierarchy>`),
  ];
  const cache = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition ${NS} saveData="0" refreshOnLoad="1" refreshedBy="Hub" refreshedDate="${(Date.now() / 86_400_000 + 25_569).toFixed(5)}" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="0" supportSubquery="1" supportAdvancedDrill="1"><cacheSource type="external" connectionId="${n}"/><cacheFields count="${cacheFields.length}">${cacheFields.join("")}</cacheFields><cacheHierarchies count="${cacheHierarchies.length}">${cacheHierarchies.join("")}</cacheHierarchies><kpis count="0"/></pivotCacheDefinition>`;

  const pivotFields = [
    ...fieldKeys.map((k) => {
      const axis = rowKeys.includes(k) ? "axisRow" : colKeys.includes(k) ? "axisCol" : "axisPage";
      const list = membersOf(k);
      return `<pivotField axis="${axis}" allDrilled="1" showAll="0" dataSourceSort="1" defaultAttributeDrillState="1"${axis === "axisPage" && list.length > 1 ? ' multipleItemSelectionAllowed="1"' : ""}><items count="${list.length + 1}">${list.map((_, i) => `<item s="1" x="${i}"/>`).join("")}<item t="default"/></items></pivotField>`;
    }),
    ...measureKeys.map(() => '<pivotField dataField="1" showAll="0"/>'),
  ];
  // Several measures: their names make a column field ("Values", index -2).
  const cols = [...colKeys.map(fieldNo), ...(measureKeys.length > 1 ? [-2] : [])];
  const pages = pageKeys.map((k) => {
    const h = olap.hierarchies.get(k)!;
    const list = membersOf(k);
    return list.length === 1 ? `<pageField fld="${fieldNo(k)}" hier="${fieldNo(k)}" name="${esc(member(h, list[0]))}" cap="${esc(list[0])}"/>` : `<pageField fld="${fieldNo(k)}" hier="${fieldNo(k)}" name="${esc(`${h}.[All]`)}" cap="All"/>`;
  });
  // A filter on several values: its members selected on the hierarchy.
  const pivotHierarchies = [
    ...fieldKeys.map((k) => {
      const list = membersOf(k);
      return pageKeys.includes(k) && list.length > 1 ? `<pivotHierarchy multipleItemSelectionAllowed="1"><members count="${list.length}" level="1">${list.map((v) => `<member name="${esc(member(olap.hierarchies.get(k)!, v))}"/>`).join("")}</members></pivotHierarchy>` : "<pivotHierarchy/>";
    }),
    ...measureKeys.map(() => '<pivotHierarchy dragToRow="0" dragToCol="0" dragToPage="0" dragToData="1"/>'),
  ];
  const top = pages.length ? pages.length + 4 : 3; // page fields above the table, a blank row between
  const lines = rowKeys.length ? new Set(rows.map((r) => rowKeys.map((k) => r[k]).join("|"))).size : 1;
  const width = Math.max(1, rowKeys.length ? 1 : 0) + Math.max(1, measureKeys.length) * Math.max(1, colKeys.length ? new Set(rows.map((r) => colKeys.map((k) => r[k]).join("|"))).size : 1);
  const pivotTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition ${NS} name="${esc(`Pivot${n}`)}" cacheId="${n}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" fieldListSortAscending="1"><location ref="A${top}:${col(width - 1)}${top + lines + (cols.length ? 2 : 1)}" firstHeaderRow="1" firstDataRow="${cols.length ? 2 : 1}" firstDataCol="1"${pages.length ? ` rowPageCount="${pages.length}" colPageCount="1"` : ""}/><pivotFields count="${pivotFields.length}">${pivotFields.join("")}</pivotFields>${
    rowKeys.length ? `<rowFields count="${rowKeys.length}">${rowKeys.map((k) => `<field x="${fieldNo(k)}"/>`).join("")}</rowFields><rowItems count="1"><i t="grand"><x/></i></rowItems>` : ""
  }${cols.length ? `<colFields count="${cols.length}">${cols.map((x) => `<field x="${x}"/>`).join("")}</colFields>` : ""}<colItems count="1"><i/></colItems>${pages.length ? `<pageFields count="${pages.length}">${pages.join("")}</pageFields>` : ""}<dataFields count="${measureKeys.length}">${measureKeys.map((m) => `<dataField fld="${measureNo(m)}" baseField="0" baseItem="0"/>`).join("")}</dataFields><pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/><pivotHierarchies count="${pivotHierarchies.length}">${pivotHierarchies.join("")}</pivotHierarchies>${
    rowKeys.length ? `<rowHierarchiesUsage count="${rowKeys.length}">${rowKeys.map((k) => `<rowHierarchyUsage hierarchyUsage="${fieldNo(k)}"/>`).join("")}</rowHierarchiesUsage>` : ""
  }${cols.length ? `<colHierarchiesUsage count="${cols.length}">${cols.map((x) => `<colHierarchyUsage hierarchyUsage="${x}"/>`).join("")}</colHierarchiesUsage>` : ""}</pivotTableDefinition>`;
  const connectionXml = `<connection id="${n}" keepAlive="1" name="${esc(`Power BI — ${model} — ${name}`)}" description="${esc(`${name} (Hub)`)}" type="5" refreshedVersion="8" background="1"><dbPr connection="${esc(connection)}" command="Model" commandType="1"/><olapPr sendLocale="1" rowDrillCount="1000"/></connection>`;
  // The figures already read, in the PivotTable's cells (and its filters above it).
  const laid = pivot(spec, rows);
  const grid = [[laid.rowLabels.join(" / ") || "", ...laid.head.map((h) => h.label)], ...laid.lines.map((l) => [l.labels.join(" / "), ...laid.head.map((h) => l.cells.get(h.id))])];
  const filterRows = pageKeys.map((k, i) => `<row r="${i + 3}">${cell(k, `A${i + 3}`)}${cell(membersOf(k).length === 1 ? membersOf(k)[0] : "(Multiple Items)", `B${i + 3}`)}</row>`);
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><sheetData><row r="1">${cell(note ? `${name} — ${note}` : name, "A1")}</row>${filterRows.join("")}${grid.map((line, i) => `<row r="${i + top}">${line.map((v, j) => cell(v, `${col(j)}${i + top}`)).join("")}</row>`).join("")}</sheetData></worksheet>`;
  return { cache, pivotTable, connectionXml, sheet, anchor: `$A$${top}` };
}

/** The parts of one pivot (n from 1): its data sheet, table, query table, connection, cache, records, PivotTable and sheet. */
function pivotParts(n: number, { name, spec, rows, dax, connection, model, note }: PivotPart) {
  // One column field at most (the others go on rows): its items are written out, as Excel wants them.
  const rowKeys = [...spec.rows, ...spec.columns.slice(1)].map(resultName);
  const colKeys = spec.columns.slice(0, 1).map(resultName);
  const keys = rows.length ? Object.keys(rows[0]) : [...rowKeys, ...colKeys, ...spec.values.map(resultName)];
  const valueKeys = keys.filter((k) => !rowKeys.includes(k) && !colKeys.includes(k));
  // The names the model gives the query's columns ('T'[col] → T[col], a value → [name]): the table keeps its own.
  const source = (k: string) => {
    const field = [...spec.rows, ...spec.columns].find((f) => resultName(f) === k);
    return field ? field.replace(/^'(.*)'(\[.*\])$/, "$1$2") : `[${k}]`;
  };
  const last = `${col(keys.length - 1)}${rows.length + 1}`;
  const table = `Data${n}`;

  const dataSheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><dimension ref="A1:${last}"/><sheetData>${[keys, ...rows.map((r) => keys.map((k) => r[k]))]
    .map((line, i) => `<row r="${i + 1}">${line.map((v, j) => cell(i ? v : String(v), `${col(j)}${i + 1}`)).join("")}</row>`)
    .join("")}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`;
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table ${NS} id="${n}" name="${table}" displayName="${table}" ref="A1:${last}" tableType="queryTable" totalsRowShown="0"><autoFilter ref="A1:${last}"/><tableColumns count="${keys.length}">${keys.map((k, i) => `<tableColumn id="${i + 1}" uniqueName="${i + 1}" name="${esc(k)}" queryTableFieldId="${i + 1}"/>`).join("")}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
  const queryTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable ${NS} name="ExternalData_${n}" connectionId="${n}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0"><queryTableRefresh nextId="${keys.length + 1}"><queryTableFields count="${keys.length}">${keys.map((k, i) => `<queryTableField id="${i + 1}" name="${esc(source(k))}" tableColumnId="${i + 1}"/>`).join("")}</queryTableFields></queryTableRefresh></queryTable>`;
  const connectionXml = `<connection id="${n}" keepAlive="1" name="${esc(`Power BI — ${model} — ${name}`)}" description="${esc(`${name} (Hub)`)}" type="5" refreshedVersion="8" background="1" saveData="1"><dbPr connection="${esc(connection)}" command="${esc(dax)}" commandType="4"/><olapPr sendLocale="1" rowDrillCount="1000"/></connection>`;

  /* Pivot: its cache (the table's fields and rows) and the table itself. */
  const axis = new Set([...rowKeys, ...colKeys]);
  const items = new Map(keys.filter((k) => axis.has(k)).map((k) => [k, [...new Set(rows.map((r) => r[k]))]]));
  const sharedItems = (k: string) => {
    const list = items.get(k);
    const values = rows.map((r) => r[k]);
    const blank = values.some((v) => v === null || v === undefined || v === "");
    if (!list) {
      const numbers = values.filter(isNumber) as number[];
      return `<sharedItems${blank ? "" : ' containsSemiMixedTypes="0"'} containsString="0" containsNumber="1"${blank ? ' containsBlank="1"' : ""}${numbers.length ? ` minValue="${Math.min(...numbers)}" maxValue="${Math.max(...numbers)}"` : ""}/>`;
    }
    const numeric = list.every((v) => isNumber(v) || v === null || v === undefined);
    const item = (v: unknown) => (v === null || v === undefined || v === "" ? "<m/>" : numeric ? `<n v="${v}"/>` : `<s v="${esc(v)}"/>`);
    return `<sharedItems${numeric ? `${blank ? "" : ' containsSemiMixedTypes="0"'} containsString="0" containsNumber="1"` : ""}${blank ? ' containsBlank="1"' : ""} count="${list.length}">${list.map(item).join("")}</sharedItems>`;
  };
  const cache = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition ${NS} r:id="rId1" refreshOnLoad="1" refreshedBy="Hub" refreshedDate="${(Date.now() / 86_400_000 + 25_569).toFixed(5)}" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${rows.length}"><cacheSource type="worksheet"><worksheetSource name="${table}"/></cacheSource><cacheFields count="${keys.length}">${keys.map((k) => `<cacheField name="${esc(k)}" numFmtId="0">${sharedItems(k)}</cacheField>`).join("")}</cacheFields></pivotCacheDefinition>`;
  const records = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords ${NS} count="${rows.length}">${rows
    .map((r) => `<r>${keys.map((k) => (items.get(k) ? `<x v="${items.get(k)!.indexOf(r[k])}"/>` : isNumber(r[k]) ? `<n v="${r[k]}"/>` : "<m/>")).join("")}</r>`)
    .join("")}</pivotCacheRecords>`;
  const fieldIndex = (k: string) => keys.indexOf(k);
  const pivotField = (k: string) => {
    const list = items.get(k);
    const on = rowKeys.includes(k) ? "axisRow" : colKeys.includes(k) ? "axisCol" : undefined;
    if (on && list) return `<pivotField axis="${on}" showAll="0"><items count="${list.length + 1}">${list.map((_, i) => `<item x="${i}"/>`).join("")}<item t="default"/></items></pivotField>`;
    return `<pivotField${valueKeys.includes(k) ? ' dataField="1"' : ""} showAll="0"/>`;
  };
  // Several values: their names make a column field ("Values", index -2).
  const cols = [...colKeys.map(fieldIndex), ...(valueKeys.length > 1 ? [-2] : [])];
  // A column field with one value: one column per item, then the total (with several values Excel lays them out itself).
  const colList = colKeys.length && valueKeys.length === 1 ? (items.get(colKeys[0]) ?? []) : [];
  const colItems = colList.length ? `<colItems count="${colList.length + 1}">${colList.map((_, i) => `<i>${i ? `<x v="${i}"/>` : "<x/>"}</i>`).join("")}<i t="grand"><x/></i></colItems>` : '<colItems count="1"><i/></colItems>';
  const pivotTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition ${NS} name="${esc(`Pivot${n}`)}" cacheId="${n}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A3:${col(Math.max(1, rowKeys.length + valueKeys.length))}${Math.max(5, (items.get(rowKeys[0])?.length ?? 1) + 4)}" firstHeaderRow="1" firstDataRow="${cols.length ? 2 : 1}" firstDataCol="1"/><pivotFields count="${keys.length}">${keys.map(pivotField).join("")}</pivotFields>${
    rowKeys.length ? `<rowFields count="${rowKeys.length}">${rowKeys.map((k) => `<field x="${fieldIndex(k)}"/>`).join("")}</rowFields><rowItems count="1"><i t="grand"><x/></i></rowItems>` : ""
  }${cols.length ? `<colFields count="${cols.length}">${cols.map((x) => `<field x="${x}"/>`).join("")}</colFields>` : ""}${colItems}<dataFields count="${valueKeys.length}">${valueKeys.map((k) => `<dataField name="${esc(`${k} `)}" fld="${fieldIndex(k)}" baseField="0" baseItem="0"/>`).join("")}</dataFields><pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
  // The pivot's figures written in its cells too: shown at once, before Excel refreshes the table.
  const laid = pivot({ ...spec, rows: rowKeys, columns: colKeys }, rows);
  const grid = [[laid.rowLabels.join(" / "), ...laid.head.map((h) => h.label)], ...laid.lines.map((l) => [l.labels.join(" / "), ...laid.head.map((h) => l.cells.get(h.id))])];
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><sheetData><row r="1">${cell(note ? `${name} — ${note}` : name, "A1")}</row>${grid.map((line, i) => `<row r="${i + 3}">${line.map((v, j) => cell(v, `${col(j)}${i + 3}`)).join("")}</row>`).join("")}</sheetData></worksheet>`;
  return { dataSheet, tableXml, queryTable, connectionXml, cache, records, pivotTable, sheet, table, last: `$${col(keys.length - 1)}$${rows.length + 1}` };
}

/** Where a pivot's PivotTable starts, and its model names when it is on the model (GETPIVOTDATA's arguments). */
type Anchor = { anchor: string; olap?: Olap };

/** The arguments of a call, split at the top-level separators (strings and parentheses kept whole). */
function splitArgs(text: string) {
  const args: string[] = [];
  let depth = 0;
  let inString = false;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') inString = !inString;
    else if (inString) continue;
    else if (c === "(") depth++;
    else if (c === ")") depth--;
    else if ((c === "," || c === ";") && depth === 0) {
      args.push(text.slice(start, i).trim());
      start = i + 1;
    }
  }
  return [...args, text.slice(start).trim()];
}

/** The end of the call starting at `open` (its "("): the index of its ")". */
function closing(text: string, open: number) {
  let depth = 0;
  let inString = false;
  for (let i = open; i < text.length; i++) {
    const c = text[i];
    if (c === '"') inString = !inString;
    else if (inString) continue;
    else if (c === "(") depth++;
    else if (c === ")" && --depth === 0) return i;
  }
  return -1;
}

const literal = (arg: string) => arg.match(/^"((?:[^"]|"")*)"$/)?.[1].replace(/""/g, '"');
const quote = (s: string) => `"${s.replace(/"/g, '""')}"`;

/**
 * A calc sheet's formula for Excel: GETPIVOTDATA's sheet name becomes the first cell of that sheet's
 * PivotTable; on a pivot on the model, the value and fields become the model's names and each item its
 * member ("[Zone].[ZoneCode].&[" & A2 & "]"), as Excel writes it for an OLAP PivotTable. ";" become ",".
 */
export function excelFormula(formula: string, pivots: Map<string, Anchor> = new Map()) {
  let out = "";
  let rest = formula;
  for (let m = rest.match(/GETPIVOTDATA\s*\(/i); m; m = rest.match(/GETPIVOTDATA\s*\(/i)) {
    const open = m.index! + m[0].length - 1;
    const end = closing(rest, open);
    if (end < 0) break;
    const args = splitArgs(rest.slice(open + 1, end)).map((a) => excelFormula(a, pivots));
    const sheet = literal(args[1] ?? "");
    const target = sheet !== undefined ? pivots.get(sheet) : undefined;
    if (target && sheet !== undefined) {
      const value = literal(args[0]) ?? "";
      const measure = target.olap && [...target.olap.measures].find(([k]) => k.toLowerCase() === value.toLowerCase())?.[1];
      const pairs: string[] = [];
      for (let i = 2; i + 1 < args.length; i += 2) {
        const field = literal(args[i]) ?? "";
        const hierarchy = target.olap && [...target.olap.hierarchies].find(([k]) => k.toLowerCase() === resultName(field).toLowerCase())?.[1];
        if (!hierarchy) pairs.push(args[i], args[i + 1]);
        else {
          const item = literal(args[i + 1]);
          pairs.push(quote(hierarchy), item !== undefined ? quote(`${hierarchy}.&[${item.replace(/\]/g, "]]")}]`) : `${quote(`${hierarchy}.&[`)}&${args[i + 1]}&"]"`);
        }
      }
      args.splice(0, args.length, measure ? quote(measure) : args[0], `'${sheet.replace(/'/g, "''")}'!${target.anchor}`, ...pairs);
    }
    out += `${rest.slice(0, m.index)}GETPIVOTDATA(${args.join(",")})`;
    rest = rest.slice(end + 1);
  }
  return (out + rest).replace(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/g, ",");
}

/** A calculation sheet: its cells, formulas with their computed value (shown before Excel recalculates). */
function calcSheetXml(sheet: CalcSheet, results: Result[][] | undefined, pivots: Map<string, Anchor>) {
  const rows = sheet.cells.map((line, r) =>
    `<row r="${r + 1}">${line
      .map((raw, c) => {
        const ref = `${col(c)}${r + 1}`;
        if (!raw.trim().startsWith("=")) return cell(/^-?\d+(\.\d+)?$/.test(raw.trim()) ? Number(raw) : raw, ref);
        const v = results?.[r]?.[c]?.value;
        const f = `<f>${esc(excelFormula(raw.trim().slice(1), pivots))}</f>`;
        return isNumber(v) ? `<c r="${ref}">${f}<v>${v}</v></c>` : typeof v === "boolean" ? `<c r="${ref}" t="b">${f}<v>${v ? 1 : 0}</v></c>` : `<c r="${ref}" t="str">${f}<v>${esc(v ?? "")}</v></c>`;
      })
      .join("")}</row>`,
  );
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><sheetData>${rows.join("")}${sheet.note ? `<row r="${sheet.cells.length + 2}">${cell(sheet.note, `A${sheet.cells.length + 2}`)}</row>` : ""}</sheetData></worksheet>`;
}

const rel = (items: [string, string, string][]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join("")}</Relationships>`;

/**
 * The workbook: the sheets in the research's order — pivots (their PivotTable; their connected data in
 * "… data" sheets, at the end) and calculations (formulas with their values, `results` by sheet name).
 */
export async function researchWorkbook(sheets: ({ kind: "pivot"; part: PivotPart } | { kind: "calc"; sheet: CalcSheet })[], results: Map<string, Result[][]>) {
  const zip = new JSZip();
  const ct = "application/vnd.openxmlformats-officedocument.spreadsheetml";
  const overrides: string[] = [`<Override PartName="/xl/workbook.xml" ContentType="${ct}.sheet.main+xml"/>`, `<Override PartName="/xl/styles.xml" ContentType="${ct}.styles+xml"/>`];
  const workbookRels: [string, string, string][] = [];
  const books: string[] = [];
  const defined: string[] = [];
  const caches: string[] = [];
  const connections: string[] = [];
  let sheetNo = 0;
  const addSheet = (name: string, xml: string, rels?: [string, string, string][]) => {
    sheetNo++;
    zip.file(`xl/worksheets/sheet${sheetNo}.xml`, xml);
    if (rels) zip.file(`xl/worksheets/_rels/sheet${sheetNo}.xml.rels`, rel(rels));
    overrides.push(`<Override PartName="/xl/worksheets/sheet${sheetNo}.xml" ContentType="${ct}.worksheet+xml"/>`);
    workbookRels.push([`rIdS${sheetNo}`, "worksheet", `worksheets/sheet${sheetNo}.xml`]);
    books.push(`<sheet name="${esc(name.slice(0, 31))}" sheetId="${sheetNo}" r:id="rIdS${sheetNo}"/>`);
    return sheetNo;
  };
  const pivots = sheets.flatMap((s, i) => (s.kind === "pivot" ? [{ n: sheets.slice(0, i + 1).filter((x) => x.kind === "pivot").length, part: s.part }] : []));
  // On the model when Excel can build it there; otherwise on a table of its query.
  const olaps = new Map(pivots.map(({ n, part }) => [n, olapOf(part.spec, part.info)]));
  const onModel = new Map(pivots.flatMap(({ n, part }) => (olaps.get(n) ? [[n, olapParts(n, part, olaps.get(n)!)]] : [])));
  const onTable = new Map(pivots.flatMap(({ n, part }) => (olaps.get(n) ? [] : [[n, pivotParts(n, part)]])));
  const anchors = new Map<string, Anchor>(pivots.map(({ n, part }) => [part.name, { anchor: onModel.get(n)?.anchor ?? "$A$3", olap: olaps.get(n) }]));
  // The visible sheets first, in order.
  for (const s of sheets) {
    if (s.kind === "calc") {
      addSheet(s.sheet.name, calcSheetXml(s.sheet, results.get(s.sheet.name), anchors));
      continue;
    }
    const n = pivots.find((p) => p.part === s.part)!.n;
    const p = onModel.get(n) ?? onTable.get(n)!;
    addSheet(s.part.name, p.sheet, [["rId1", "pivotTable", `../pivotTables/pivotTable${n}.xml`]]);
    zip.file(`xl/pivotTables/pivotTable${n}.xml`, p.pivotTable);
    zip.file(`xl/pivotTables/_rels/pivotTable${n}.xml.rels`, rel([["rId1", "pivotCacheDefinition", `../pivotCache/pivotCacheDefinition${n}.xml`]]));
    zip.file(`xl/pivotCache/pivotCacheDefinition${n}.xml`, p.cache);
    workbookRels.push([`rIdC${n}`, "pivotCacheDefinition", `pivotCache/pivotCacheDefinition${n}.xml`]);
    caches.push(`<pivotCache cacheId="${n}" r:id="rIdC${n}"/>`);
    overrides.push(`<Override PartName="/xl/pivotTables/pivotTable${n}.xml" ContentType="${ct}.pivotTable+xml"/>`, `<Override PartName="/xl/pivotCache/pivotCacheDefinition${n}.xml" ContentType="${ct}.pivotCacheDefinition+xml"/>`);
    if (onModel.has(n)) {
      connections.push(onModel.get(n)!.connectionXml);
      continue;
    }
    zip.file(`xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`, rel([["rId1", "pivotCacheRecords", `pivotCacheRecords${n}.xml`]]));
    zip.file(`xl/pivotCache/pivotCacheRecords${n}.xml`, onTable.get(n)!.records);
    overrides.push(`<Override PartName="/xl/pivotCache/pivotCacheRecords${n}.xml" ContentType="${ct}.pivotCacheRecords+xml"/>`);
  }
  // Then the connected data of each pivot on a table.
  for (const { n, part } of pivots) {
    const p = onTable.get(n);
    if (!p) continue;
    const index = addSheet(`${part.name.slice(0, 25)} data`, p.dataSheet, [["rId1", "table", `../tables/table${n}.xml`]]);
    zip.file(`xl/tables/table${n}.xml`, p.tableXml);
    zip.file(`xl/tables/_rels/table${n}.xml.rels`, rel([["rId1", "queryTable", `../queryTables/queryTable${n}.xml`]]));
    zip.file(`xl/queryTables/queryTable${n}.xml`, p.queryTable);
    overrides.push(`<Override PartName="/xl/tables/table${n}.xml" ContentType="${ct}.table+xml"/>`, `<Override PartName="/xl/queryTables/queryTable${n}.xml" ContentType="${ct}.queryTable+xml"/>`);
    connections.push(p.connectionXml);
    defined.push(`<definedName name="ExternalData_${n}" localSheetId="${index - 1}" hidden="1">'${esc(`${part.name.slice(0, 25)} data`.replace(/'/g, "''"))}'!$A$1:${p.last}</definedName>`);
  }
  if (connections.length) {
    zip.file("xl/connections.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections ${NS}>${connections.sort((a, b) => Number(a.match(/id="(\d+)"/)![1]) - Number(b.match(/id="(\d+)"/)![1])).join("")}</connections>`);
    overrides.push(`<Override PartName="/xl/connections.xml" ContentType="${ct}.connections+xml"/>`);
    workbookRels.push(["rIdX", "connections", "connections.xml"]);
  }
  workbookRels.push(["rIdY", "styles", "styles.xml"]);
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join("")}</Types>`,
  );
  zip.file("_rels/.rels", rel([["rId1", "officeDocument", "xl/workbook.xml"]]));
  // Formulas computed again when the file opens (fullCalcOnLoad): GETPIVOTDATA reads the refreshed pivots.
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook ${NS}><bookViews><workbookView activeTab="0"/></bookViews><sheets>${books.join("")}</sheets>${defined.length ? `<definedNames>${defined.join("")}</definedNames>` : ""}<calcPr calcId="191029" fullCalcOnLoad="1"/>${caches.length ? `<pivotCaches>${caches.join("")}</pivotCaches>` : ""}</workbook>`,
  );
  zip.file("xl/_rels/workbook.xml.rels", rel(workbookRels));
  zip.file(
    "xl/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  );
  return zip.generateAsync({ type: "blob", mimeType: `${ct}.sheet` });
}
