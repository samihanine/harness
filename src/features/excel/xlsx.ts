/**
 * What the Graph workbook API cannot write, written into the .xlsx itself (lib/ooxml): for each
 * option column, a drop-down list of its options (data validation) and one colored fill per option
 * (conditional formatting), so the file shows the app's colors and choices in Excel too. The file is
 * read, changed and sent back only if nobody saved it meanwhile (its version).
 */
import { EXCEL_FILLS } from "@/components/badge";
import { OoxmlPackage, REL, element, kid, kids, resolvePart, type XmlDoc, type XmlEl } from "@/lib/ooxml/package";
import { readItem, writeItem } from "@/services/ms";
import { columnIndex, letter, type ExcelRef } from "@/services/excel";
import type { Column } from "./table";

/** The order of a worksheet's elements (the schema requires it). */
const SHEET_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
const MARK = "Hub";
const LAST_ROW = 1_048_576;

function insertInOrder(sheet: XmlEl, el: XmlEl) {
  const rank = SHEET_ORDER.indexOf(el.nodeName);
  const after = kids(sheet).find((c) => SHEET_ORDER.indexOf(c.nodeName) > rank);
  sheet.insertBefore(el, after ?? null);
}

/** The worksheet part holding a table, and the table's range. */
async function sheetOf(pkg: OoxmlPackage, table: string) {
  for (const part of pkg.parts("xl/worksheets/").filter((p) => /sheet\d*\.xml$/.test(p))) {
    for (const rel of (await pkg.rels(part)).filter((r) => r.type === REL.table)) {
      const t = (await pkg.xml(resolvePart(part, rel.target))).documentElement as XmlEl;
      if (t.getAttribute("name") === table || t.getAttribute("displayName") === table) {
        const ref = t.getAttribute("ref")!;
        const [, c1, r1] = ref.match(/^([A-Z]+)(\d+)/)!;
        const names = kids(kid(t, "tableColumns")!, "tableColumn").map((c) => c.getAttribute("name") ?? "");
        return { part, left: columnIndex(c1), top: Number(r1), names };
      }
    }
  }
  throw new Error(`The table "${table}" was not found in the file.`);
}

/** A differential style (fill + font color) for conditional formatting: an identical one is reused. */
function dxfOf(styles: XmlDoc, fill: string, font: string) {
  const root = styles.documentElement as XmlEl;
  let dxfs = kid(root, "dxfs");
  if (!dxfs) {
    dxfs = element(styles, "dxfs", { count: 0 });
    const after = kids(root).find((c) => ["tableStyles", "colors", "extLst"].includes(c.nodeName));
    root.insertBefore(dxfs, after ?? null);
  }
  const argb = (hex: string) => `FF${hex.replace("#", "").toUpperCase()}`;
  const list = kids(dxfs, "dxf");
  const same = list.findIndex((d) => d.getElementsByTagName("bgColor")[0]?.getAttribute("rgb") === argb(fill) && d.getElementsByTagName("color")[0]?.getAttribute("rgb") === argb(font));
  if (same >= 0) return same;
  const ns = root.namespaceURI;
  const make = (name: string, attrs: Record<string, string> = {}, children: XmlEl[] = []) => {
    const el = styles.createElementNS(ns, name);
    Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
    children.forEach((c) => el.appendChild(c));
    return el;
  };
  dxfs.appendChild(make("dxf", {}, [make("font", {}, [make("color", { rgb: argb(font) })]), make("fill", {}, [make("patternFill", { patternType: "solid" }, [make("bgColor", { rgb: argb(fill) })])])]));
  dxfs.setAttribute("count", String(list.length + 1));
  return list.length;
}

/**
 * Writes the option columns' lists and colors into the file (the app's earlier ones replaced, the
 * user's own rules on other cells kept). `header`: the file's columns as the app expects them (a
 * file not yet showing the last column changes is refused: try again in a moment).
 */
export async function writeChoices(e: ExcelRef, columns: Column[], header: string[]) {
  const { data, eTag } = await readItem(e);
  const pkg = await OoxmlPackage.open(data);
  await applyChoices(pkg, e.table, columns, header);
  await writeItem({ driveId: e.driveId, itemId: e.itemId, data: await pkg.save("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), eTag });
}

/** The same on an opened file (the tests use it directly). */
export async function applyChoices(pkg: OoxmlPackage, table: string, columns: Column[], header: string[]) {
  const { part, left, top, names } = await sheetOf(pkg, table);
  if (JSON.stringify(names) !== JSON.stringify(header)) throw new Error("NOT_READY");
  const doc = await pkg.xml(part);
  const sheet = doc.documentElement as XmlEl;
  const ns = sheet.namespaceURI;
  const make = (name: string, attrs: Record<string, string | number> = {}, text?: string) => {
    const el = doc.createElementNS(ns, name);
    Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, String(v)));
    if (text !== undefined) el.appendChild(doc.createTextNode(text));
    return el;
  };
  const ranges = new Set(header.map((_, i) => `${letter(left + i)}${top + 1}:${letter(left + i)}${LAST_ROW}`));
  // The app's earlier rules on these columns go (validation marked "Hub"; formatting on a whole column of the table).
  for (const cf of kids(sheet, "conditionalFormatting")) if (ranges.has(cf.getAttribute("sqref") ?? "")) sheet.removeChild(cf);
  let validations = kid(sheet, "dataValidations");
  for (const dv of validations ? kids(validations, "dataValidation") : []) if (dv.getAttribute("promptTitle") === MARK || ranges.has(dv.getAttribute("sqref") ?? "")) validations!.removeChild(dv);

  const styles = pkg.has("xl/styles.xml") ? await pkg.xml("xl/styles.xml") : undefined;
  let priority = Math.max(0, ...Array.from(sheet.getElementsByTagName("cfRule") as ArrayLike<XmlEl>).map((r) => Number(r.getAttribute("priority")) || 0));
  for (const [i, name] of header.entries()) {
    const column = columns.find((c) => c.name === name);
    if (column?.type !== "option" || !column.options?.length) continue;
    const sqref = `${letter(left + i)}${top + 1}:${letter(left + i)}${LAST_ROW}`;
    // A list in the cell (its options; Excel allows 255 characters).
    const list = column.options.map((o) => o.value.replace(/,/g, " ")).join(",");
    if (list.length <= 255) {
      if (!validations) insertInOrder(sheet, (validations = make("dataValidations", { count: 0 })));
      const dv = make("dataValidation", { type: "list", allowBlank: 1, showErrorMessage: 1, errorStyle: "warning", promptTitle: MARK, sqref });
      dv.appendChild(make("formula1", {}, `"${list.replace(/"/g, '""')}"`));
      validations.appendChild(dv);
    }
    // A fill per option.
    if (styles) {
      const cf = make("conditionalFormatting", { sqref });
      for (const o of column.options) {
        const { fill, font } = EXCEL_FILLS[o.color];
        const rule = make("cfRule", { type: "cellIs", dxfId: dxfOf(styles, fill, font), priority: ++priority, operator: "equal" });
        rule.appendChild(make("formula", {}, `"${o.value.replace(/"/g, '""')}"`));
        cf.appendChild(rule);
      }
      insertInOrder(sheet, cf);
    }
  }
  if (validations) {
    validations.setAttribute("count", String(kids(validations, "dataValidation").length));
    if (!kids(validations, "dataValidation").length) sheet.removeChild(validations);
  }
}

/** The same, waiting for Excel for the web to save the last column changes into the file first. */
export async function writeChoicesWhenSaved(e: ExcelRef, columns: Column[], header: string[]) {
  for (let i = 0; ; i++) {
    try {
      return await writeChoices(e, columns, header);
    } catch (error) {
      if (!(error instanceof Error && error.message === "NOT_READY") || i >= 8) throw error instanceof Error && error.message === "NOT_READY" ? new Error("Excel has not saved the new columns into the file yet: sync again in a minute.") : error;
      await new Promise((r) => setTimeout(r, 3000));
    }
  }
}
