/**
 * What Excel's API (Graph v1.0) cannot do, written in the workbook's file itself: conditional formats and
 * drop-down lists, as Excel writes them (a sheet's <conditionalFormatting> / <dataValidations>, the formats in
 * styles.xml's <dxfs>). Only these parts change: everything else of the file (charts, connections, tables) stays.
 */
import JSZip from "jszip";
import { closeSession, replaceWorkbook, workbookBytes } from "./excel";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const argb = (c: string) => `FF${c.replace("#", "").toUpperCase().padStart(6, "0").slice(-6)}`;

/** The file of a sheet (xl/worksheets/sheetN.xml) by its name. */
async function sheetFile(zip: JSZip, name: string) {
  const wb = await zip.file("xl/workbook.xml")!.async("string");
  const sheet = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => m[0]).find((t) => t.match(/name="([^"]*)"/)?.[1]?.replace(/&amp;/g, "&").replace(/&apos;/g, "'") === name);
  if (!sheet) throw new Error(`No sheet "${name}".`);
  const rid = sheet.match(/r:id="([^"]+)"/)![1];
  const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  const target = rels.match(new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"`))?.[1] ?? rels.match(new RegExp(`<Relationship[^>]*Target="([^"]+)"[^>]*Id="${rid}"`))?.[1];
  return target!.startsWith("/") ? target!.slice(1) : `xl/${target}`;
}

/** Puts an element in a sheet at its place in the schema's order (before the first element that must follow it). */
function insert(xml: string, element: string, after: string[]) {
  const follow = ["dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"].filter((t) => !after.includes(t));
  for (const tag of follow) {
    const at = xml.search(new RegExp(`<${tag}[\\s>/]`));
    if (at >= 0) return xml.slice(0, at) + element + xml.slice(at);
  }
  return xml.replace("</worksheet>", `${element}</worksheet>`);
}

/** A differential format (the look a rule gives) added to styles.xml; returns its index. */
async function addDxf(zip: JSZip, f: { fill?: string; fontColor?: string; bold?: boolean }) {
  let styles = await zip.file("xl/styles.xml")!.async("string");
  const dxf = `<dxf>${f.fontColor || f.bold ? `<font>${f.bold ? "<b/>" : ""}${f.fontColor ? `<color rgb="${argb(f.fontColor)}"/>` : ""}</font>` : ""}${f.fill ? `<fill><patternFill patternType="solid"><fgColor rgb="${argb(f.fill)}"/><bgColor rgb="${argb(f.fill)}"/></patternFill></fill>` : ""}</dxf>`;
  let index: number;
  if (/<dxfs[^>]*\/>/.test(styles)) {
    styles = styles.replace(/<dxfs[^>]*\/>/, `<dxfs count="1">${dxf}</dxfs>`);
    index = 0;
  } else if (/<dxfs/.test(styles)) {
    index = (styles.match(/<dxfs[^>]*>([\s\S]*?)<\/dxfs>/)![1].match(/<dxf>/g) ?? []).length;
    styles = styles.replace(/<dxfs[^>]*>([\s\S]*?)<\/dxfs>/, (_, inner) => `<dxfs count="${index + 1}">${inner}${dxf}</dxfs>`);
  } else {
    index = 0;
    const at = styles.search(/<(tableStyles|colors|extLst)[\s>/]/);
    styles = at >= 0 ? styles.slice(0, at) + `<dxfs count="1">${dxf}</dxfs>` + styles.slice(at) : styles.replace("</styleSheet>", `<dxfs count="1">${dxf}</dxfs></styleSheet>`);
  }
  zip.file("xl/styles.xml", styles);
  return index;
}

async function edit(itemId: string, change: (zip: JSZip) => Promise<void>) {
  await closeSession(itemId);
  const zip = await JSZip.loadAsync(await workbookBytes(itemId));
  await change(zip);
  await replaceWorkbook(itemId, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
}

export type Rule =
  | { rule: "cellValue"; operator: string; formula1: string; formula2?: string }
  | { rule: "text"; text: string }
  | { rule: "custom"; formula: string }
  | { rule: "colorScale"; colors: string[] }
  | { rule: "dataBar"; fill?: string };

const OPERATORS: Record<string, string> = { greaterthan: "greaterThan", lessthan: "lessThan", between: "between", notbetween: "notBetween", equalto: "equal", equal: "equal", notequalto: "notEqual", notequal: "notEqual", greaterthanorequal: "greaterThanOrEqual", lessthanorequal: "lessThanOrEqual" };

/** A conditional format on a range of a sheet (`clear`: the range's other rules removed first). */
export async function addConditionalFormat(itemId: string, sheet: string, range: string, r: Rule, look: { fill?: string; fontColor?: string; bold?: boolean }, clear = false) {
  await edit(itemId, async (zip) => {
    const file = await sheetFile(zip, sheet);
    let xml = await zip.file(file)!.async("string");
    const sqref = range.replace(/\$/g, "").replace(/,/g, " ");
    if (clear) xml = xml.replace(new RegExp(`<conditionalFormatting sqref="${sqref.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[\\s\\S]*?</conditionalFormatting>`, "g"), "");
    const priority = 1 + Math.max(0, ...[...xml.matchAll(/priority="(\d+)"/g)].map((m) => Number(m[1])));
    const first = sqref.split(/[\s:]/)[0];
    const formula = (f: string) => `<formula>${esc(f.replace(/^=/, ""))}</formula>`;
    let rule: string;
    if (r.rule === "colorScale") {
      const c = r.colors.length ? r.colors : ["#F8696B", "#FFEB84", "#63BE7B"];
      rule = `<cfRule type="colorScale" priority="${priority}"><colorScale><cfvo type="min"/>${c.length >= 3 ? '<cfvo type="percentile" val="50"/>' : ""}<cfvo type="max"/>${c.slice(0, 3).map((x) => `<color rgb="${argb(x)}"/>`).join("")}</colorScale></cfRule>`;
    } else if (r.rule === "dataBar") {
      rule = `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(r.fill ?? "#B39251")}"/></dataBar></cfRule>`;
    } else {
      const dxf = await addDxf(zip, look);
      if (r.rule === "cellValue") {
        const op = OPERATORS[r.operator.toLowerCase()] ?? "greaterThan";
        rule = `<cfRule type="cellIs" dxfId="${dxf}" priority="${priority}" operator="${op}">${formula(r.formula1)}${r.formula2 ? formula(r.formula2) : ""}</cfRule>`;
      } else if (r.rule === "text") rule = `<cfRule type="containsText" dxfId="${dxf}" priority="${priority}" operator="containsText" text="${esc(r.text)}">${formula(`NOT(ISERROR(SEARCH("${r.text.replace(/"/g, '""')}",${first})))`)}</cfRule>`;
      else rule = `<cfRule type="expression" dxfId="${dxf}" priority="${priority}">${formula(r.formula)}</cfRule>`;
    }
    xml = insert(xml, `<conditionalFormatting sqref="${sqref}">${rule}</conditionalFormatting>`, []);
    zip.file(file, xml);
  });
}

/** A drop-down list on a range (values, or a source range "=Lists!$A$2:$A$9"); none: the range's validation removed. */
export async function setValidation(itemId: string, sheet: string, range: string, list?: { values?: string[]; source?: string; message?: string }) {
  await edit(itemId, async (zip) => {
    const file = await sheetFile(zip, sheet);
    let xml = await zip.file(file)!.async("string");
    const sqref = range.replace(/\$/g, "").replace(/,/g, " ");
    const existing = xml.match(/<dataValidations[^>]*>([\s\S]*?)<\/dataValidations>/)?.[1] ?? "";
    const kept = [...existing.matchAll(/<dataValidation\b[\s\S]*?<\/dataValidation>/g)].map((m) => m[0]).filter((v) => !v.includes(`sqref="${sqref}"`));
    if (list) {
      const f = list.source ? list.source.replace(/^=/, "") : `"${(list.values ?? []).join(",").replace(/"/g, '""')}"`;
      kept.push(`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"${list.message ? ` prompt="${esc(list.message)}"` : ""} sqref="${sqref}"><formula1>${esc(f)}</formula1></dataValidation>`);
    }
    xml = xml.replace(/<dataValidations[^>]*>[\s\S]*?<\/dataValidations>/, "");
    if (kept.length) xml = insert(xml, `<dataValidations count="${kept.length}">${kept.join("")}</dataValidations>`, ["dataValidations"]);
    zip.file(file, xml);
  });
}
