/**
 * Rules the Graph workbook API cannot write (no conditional formats, no data validation): the choices'
 * colors and dropdowns, and the ratings' check, written into the file itself. The workbook is read,
 * only the table's sheet (and the styles of the colors) changes, and it is sent back only if it did not
 * change meanwhile (eTag). The rules cover rows added later too (e.g. a form's answers).
 */
import JSZip from "jszip";
import { readItem, writeItem } from "@/server/ms";
import type { Color, Field } from "./schema";

export type ColumnRule = { letter: string; field: Field };
export type CellColors = Record<Color, { fill: string; font: string }>;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const unesc = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
const argb = (hex: string) => `FF${hex.replace("#", "").toUpperCase()}`;
/** An Excel text literal in a formula ("" doubled). */
const text = (s: string) => `"${s.replace(/"/g, '""')}"`;

/** Elements of a worksheet that come after the conditional formats, in the file format's order. */
const AFTER_CF = ["dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
const AFTER_DV = AFTER_CF.slice(1);
/** Inserts `xml` before the first of these elements (else before the end). */
function insertBefore(sheet: string, names: string[], xml: string) {
  const at = names.map((n) => sheet.search(new RegExp(`<${n}[\\s>/]`))).filter((i) => i >= 0).sort((a, b) => a - b)[0] ?? sheet.lastIndexOf("</worksheet>");
  return sheet.slice(0, at) + xml + sheet.slice(at);
}

/** The worksheet part of a sheet name. */
async function sheetPath(zip: JSZip, name: string) {
  const workbook = await zip.file("xl/workbook.xml")!.async("string");
  const sheet = [...workbook.matchAll(/<sheet\b[^>]*>/g)].map((m) => m[0]).find((tag) => unesc(tag.match(/\bname="([^"]*)"/)?.[1] ?? "") === name);
  const rid = sheet?.match(/\br:id="([^"]*)"/)?.[1];
  if (!rid) throw new Error(`No sheet "${name}" in the workbook.`);
  const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  const target = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).find((tag) => tag.includes(`Id="${rid}"`))?.match(/Target="([^"]*)"/)?.[1];
  if (!target) throw new Error(`Sheet "${name}" not found in the workbook's parts.`);
  return target.startsWith("/") ? target.slice(1) : `xl/${target}`;
}

/** Adds one differential style per color used (fill + font of a choice) and returns their ids. */
function addColors(styles: string, colors: Color[], palette: CellColors) {
  const dxf = (c: Color) => `<dxf><font><color rgb="${argb(palette[c].font)}"/></font><fill><patternFill patternType="solid"><bgColor rgb="${argb(palette[c].fill)}"/></patternFill></fill></dxf>`;
  const added = colors.map(dxf).join("");
  const ids = new Map<Color, number>();
  const block = styles.match(/<dxfs\b[^>]*?(\/>|>[\s\S]*?<\/dxfs>)/);
  const count = block && block[1] !== "/>" ? (block[0].match(/<dxf[\s>]/g) ?? []).length : 0;
  colors.forEach((c, i) => ids.set(c, count + i));
  if (block && block[1] !== "/>") return { ids, styles: styles.replace(block[0], block[0].replace(/count="\d+"/, `count="${count + colors.length}"`).replace("</dxfs>", `${added}</dxfs>`)) };
  const fresh = `<dxfs count="${colors.length}">${added}</dxfs>`;
  if (block) return { ids, styles: styles.replace(block[0], fresh) };
  // In its place: after the cell styles, before the table styles (file format's order).
  const at = ["tableStyles", "colors", "extLst"].map((n) => styles.search(new RegExp(`<${n}[\\s>/]`))).filter((i) => i >= 0).sort((x, y) => x - y)[0] ?? styles.lastIndexOf("</styleSheet>");
  return { ids, styles: styles.slice(0, at) + fresh + styles.slice(at) };
}

/**
 * Writes the rules of the columns (rows `first`..`last` of the sheet): the colors of each choice
 * (badges' colors), a dropdown for single choices, whole numbers for ratings. The rules this app wrote
 * before on these columns are replaced; other rules of the sheet are kept.
 */
export async function writeColumnRules(file: { driveId: string; itemId: string }, sheet: string, columns: ColumnRule[], first: number, last: number, palette: CellColors) {
  const { base64, eTag } = await readItem({ data: file });
  const zip = await JSZip.loadAsync(base64, { base64: true });
  const path = await sheetPath(zip, sheet);
  let xml = await zip.file(path)!.async("string");
  const ranges = new Set(columns.map((c) => `${c.letter}${first}:${c.letter}${last}`));
  // Rules written before on these columns go (they are rewritten from the schema).
  xml = xml.replace(/<conditionalFormatting\b[^>]*sqref="([^"]*)"[^>]*>[\s\S]*?<\/conditionalFormatting>/g, (all, sqref: string) => (ranges.has(sqref) ? "" : all));
  xml = xml.replace(/<dataValidation\b[^>]*sqref="([^"]*)"[^>]*?(\/>|>[\s\S]*?<\/dataValidation>)/g, (all, sqref: string) => (ranges.has(sqref) ? "" : all));
  xml = xml.replace(/<dataValidations\b[^>]*>\s*<\/dataValidations>/g, "");
  const left = (xml.match(/<dataValidation\b/g) ?? []).length;
  xml = xml.replace(/(<dataValidations\b[^>]*count=")\d+"/, `$1${left}"`);

  const choices = columns.filter((c) => c.field.type === "option" && c.field.options?.length);
  const colors = [...new Set(choices.flatMap((c) => c.field.options!.map((o) => o.color)))];
  const stylesPath = "xl/styles.xml";
  const { styles, ids } = addColors(await zip.file(stylesPath)!.async("string"), colors, palette);
  zip.file(stylesPath, styles);

  let priority = Math.max(0, ...[...xml.matchAll(/priority="(\d+)"/g)].map((m) => Number(m[1]))) + 1;
  const cf = choices
    .map(({ letter, field }) => {
      const range = `${letter}${first}:${letter}${last}`;
      const rules = field.options!.map((o) =>
        field.multiple
          ? `<cfRule type="containsText" dxfId="${ids.get(o.color)}" priority="${priority++}" operator="containsText" text="${esc(o.value)}"><formula>${esc(`NOT(ISERROR(SEARCH(${text(o.value)},${letter}${first})))`)}</formula></cfRule>`
          : `<cfRule type="cellIs" dxfId="${ids.get(o.color)}" priority="${priority++}" operator="equal"><formula>${esc(text(o.value))}</formula></cfRule>`,
      );
      return `<conditionalFormatting sqref="${range}">${rules.join("")}</conditionalFormatting>`;
    })
    .join("");
  if (cf) xml = insertBefore(xml, AFTER_CF, cf);

  const dv = columns
    .map(({ letter, field: f }) => {
      const range = `${letter}${first}:${letter}${last}`;
      if (f.type === "rating") return `<dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" errorTitle="Invalid rating" error="${esc(`A whole number from 1 to ${f.max ?? 5}`)}" sqref="${range}"><formula1>1</formula1><formula2>${f.max ?? 5}</formula2></dataValidation>`;
      if (f.type !== "option" || f.multiple || !f.options?.length) return "";
      const list = f.options.map((o) => o.value).join(",");
      // A list of values in a rule: no commas inside a value, 255 characters at most.
      if (list.length > 255 || f.options.some((o) => o.value.includes(",") || o.value.includes('"'))) return "";
      return `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Invalid value" error="${esc(`Choose one of: ${f.options.map((o) => o.value).join(", ")}`.slice(0, 225))}" sqref="${range}"><formula1>${esc(`"${list}"`)}</formula1></dataValidation>`;
    })
    .filter(Boolean);
  if (dv.length) {
    if (/<dataValidations\b/.test(xml)) xml = xml.replace(/(<dataValidations\b[^>]*count=")(\d+)("[^>]*>)/, (_, a, n, b) => `${a}${Number(n) + dv.length}${b}${dv.join("")}`);
    else xml = insertBefore(xml, AFTER_DV, `<dataValidations count="${dv.length}">${dv.join("")}</dataValidations>`);
  }
  zip.file(path, xml);
  const out = await zip.generateAsync({ type: "base64", compression: "DEFLATE" });
  await writeItem({ data: { ...file, base64: out, eTag } });
  return { colors: choices.length, dropdowns: dv.length };
}
