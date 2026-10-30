/**
 * A pivot card as an Excel file connected to its semantic model: a "Data" sheet holding the card's
 * query as a table linked to the model (an OLE DB connection to Power BI, as "Analyze in Excel" makes,
 * running the card's DAX: Data › Refresh All reads the model again), and a "Pivot" sheet with a real
 * PivotTable on that table (rows, columns, values of the card). Written by hand (the libraries in reach
 * do not write pivot tables); the rows already read fill the file, so it opens with the figures.
 */
import JSZip from "jszip";
import type { Rows } from "@/services/pbi";
import { pivot, resultName, type PivotSpec } from "./research";

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

/**
 * The workbook. `dax`: the card's query (without the app's row limit); `rows`: what it returned;
 * `connection`: the OLE DB connection string; `name`: the model's name (shown in Excel).
 */
export async function pivotWorkbook({ title, spec, rows, dax, connection, name }: { title: string; spec: PivotSpec; rows: Rows; dax: string; connection: string; name: string }) {
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

  /* Data sheet: the query's table. */
  const cell = (v: unknown, ref: string) => (isNumber(v) ? `<c r="${ref}"><v>${v}</v></c>` : v === null || v === undefined || v === "" ? "" : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`);
  const dataSheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><dimension ref="A1:${last}"/><sheetData>${[keys, ...rows.map((r) => keys.map((k) => r[k]))]
    .map((line, i) => `<row r="${i + 1}">${line.map((v, j) => cell(i ? v : String(v), `${col(j)}${i + 1}`)).join("")}</row>`)
    .join("")}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`;
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table ${NS} id="1" name="Data" displayName="Data" ref="A1:${last}" tableType="queryTable" totalsRowShown="0"><autoFilter ref="A1:${last}"/><tableColumns count="${keys.length}">${keys.map((k, i) => `<tableColumn id="${i + 1}" uniqueName="${i + 1}" name="${esc(k)}" queryTableFieldId="${i + 1}"/>`).join("")}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
  const queryTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable ${NS} name="ExternalData_1" connectionId="1" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0"><queryTableRefresh nextId="${keys.length + 1}"><queryTableFields count="${keys.length}">${keys.map((k, i) => `<queryTableField id="${i + 1}" name="${esc(source(k))}" tableColumnId="${i + 1}"/>`).join("")}</queryTableFields></queryTableRefresh></queryTable>`;
  const connections = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections ${NS}><connection id="1" keepAlive="1" name="${esc(`Power BI — ${name}`)}" description="${esc(`${title} (Hub)`)}" type="5" refreshedVersion="8" background="1" saveData="1"><dbPr connection="${esc(connection)}" command="${esc(dax)}" commandType="4"/><olapPr sendLocale="1" rowDrillCount="1000"/></connection></connections>`;

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
<pivotCacheDefinition ${NS} r:id="rId1" refreshOnLoad="1" refreshedBy="Hub" refreshedDate="${(Date.now() / 86_400_000 + 25_569).toFixed(5)}" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${rows.length}"><cacheSource type="worksheet"><worksheetSource name="Data"/></cacheSource><cacheFields count="${keys.length}">${keys.map((k) => `<cacheField name="${esc(k)}" numFmtId="0">${sharedItems(k)}</cacheField>`).join("")}</cacheFields></pivotCacheDefinition>`;
  const records = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords ${NS} count="${rows.length}">${rows
    .map(
      (r) =>
        `<r>${keys
          .map((k) => {
            const list = items.get(k);
            if (list) return `<x v="${list.indexOf(r[k])}"/>`;
            return isNumber(r[k]) ? `<n v="${r[k]}"/>` : "<m/>";
          })
          .join("")}</r>`,
    )
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
<pivotTableDefinition ${NS} name="${esc(title.slice(0, 60) || "Pivot")}" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A3:${col(Math.max(1, rowKeys.length + valueKeys.length))}${Math.max(5, (items.get(rowKeys[0])?.length ?? 1) + 4)}" firstHeaderRow="1" firstDataRow="${cols.length ? 2 : 1}" firstDataCol="1"/><pivotFields count="${keys.length}">${keys.map(pivotField).join("")}</pivotFields>${
    rowKeys.length ? `<rowFields count="${rowKeys.length}">${rowKeys.map((k) => `<field x="${fieldIndex(k)}"/>`).join("")}</rowFields><rowItems count="1"><i t="grand"><x/></i></rowItems>` : ""
  }${cols.length ? `<colFields count="${cols.length}">${cols.map((x) => `<field x="${x}"/>`).join("")}</colFields>` : ""}${colItems}<dataFields count="${valueKeys.length}">${valueKeys.map((k) => `<dataField name="${esc(`${k} `)}" fld="${fieldIndex(k)}" baseField="0" baseItem="0"/>`).join("")}</dataFields><pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
  // The pivot's figures written in its cells too: shown at once, before Excel refreshes the table.
  const laid = pivot({ ...spec, rows: rowKeys, columns: colKeys }, rows);
  const grid = [[laid.rowLabels.join(" / "), ...laid.head.map((h) => h.label)], ...laid.lines.map((l) => [l.labels.join(" / "), ...laid.head.map((h) => l.cells.get(h.id))])];
  const pivotSheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><sheetData><row r="1">${cell(title, "A1")}</row>${grid.map((line, i) => `<row r="${i + 3}">${line.map((v, j) => cell(v, `${col(j)}${i + 3}`)).join("")}</row>`).join("")}</sheetData></worksheet>`;

  const zip = new JSZip();
  const ct = "application/vnd.openxmlformats-officedocument.spreadsheetml";
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${ct}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${ct}.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="${ct}.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="${ct}.styles+xml"/><Override PartName="/xl/connections.xml" ContentType="${ct}.connections+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="${ct}.table+xml"/><Override PartName="/xl/queryTables/queryTable1.xml" ContentType="${ct}.queryTable+xml"/><Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="${ct}.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="${ct}.pivotCacheRecords+xml"/><Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="${ct}.pivotTable+xml"/></Types>`,
  );
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook ${NS}><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="ExternalData_1" localSheetId="1" hidden="1">Data!$A$1:$${col(keys.length - 1)}$${rows.length + 1}</definedName></definedNames><pivotCaches><pivotCache cacheId="1" r:id="rId5"/></pivotCaches></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId4" Type="${REL}/connections" Target="connections.xml"/><Relationship Id="rId5" Type="${REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
  );
  zip.file(
    "xl/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  );
  zip.file("xl/connections.xml", connections);
  zip.file("xl/worksheets/sheet1.xml", pivotSheet);
  zip.file("xl/worksheets/_rels/sheet1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>`);
  zip.file("xl/worksheets/sheet2.xml", dataSheet);
  zip.file("xl/worksheets/_rels/sheet2.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`);
  zip.file("xl/tables/table1.xml", table);
  zip.file("xl/tables/_rels/table1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>`);
  zip.file("xl/queryTables/queryTable1.xml", queryTable);
  zip.file("xl/pivotCache/pivotCacheDefinition1.xml", cache);
  zip.file("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>`);
  zip.file("xl/pivotCache/pivotCacheRecords1.xml", records);
  zip.file("xl/pivotTables/pivotTable1.xml", pivotTable);
  zip.file("xl/pivotTables/_rels/pivotTable1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`);
  return zip.generateAsync({ type: "blob", mimeType: `${ct}.sheet` });
}
