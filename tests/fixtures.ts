/** Decks made for the tests (and the agent evals): plain text boxes, a title, a picture, notes. */
import PptxGenJS from "pptxgenjs";

const PIXEL = "image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** A 3-slide deck: a cover, a figures slide, a closing slide. */
export async function sampleDeck() {
  const p = new PptxGenJS();
  p.layout = "LAYOUT_WIDE";
  const cover = p.addSlide();
  cover.addText("Quarterly review Q2", { x: 0.5, y: 2, w: 12, h: 1.5, fontSize: 40, bold: true, objectName: "Title" });
  cover.addText("Finance team", { x: 0.5, y: 3.6, w: 12, h: 0.6, fontSize: 20, objectName: "Subtitle" });
  const figures = p.addSlide();
  figures.addText("Results Q2", { x: 0.5, y: 0.3, w: 12, h: 0.9, fontSize: 32, objectName: "Heading" });
  figures.addText("Revenue 11.2 M€\nMargin 36 %", { x: 0.5, y: 1.5, w: 6, h: 3, fontSize: 18, bullet: true, objectName: "Body" });
  figures.addImage({ data: PIXEL, x: 7, y: 1.5, w: 5, h: 3, objectName: "Chart picture" });
  figures.addNotes("Figures from the finance model.");
  const end = p.addSlide();
  end.addText("Thank you", { x: 0.5, y: 3, w: 12, h: 1, fontSize: 36, objectName: "Closing" });
  return (await p.write({ outputType: "uint8array" })) as Uint8Array;
}

/** A template deck: each slide a type, named in its notes. */
export async function templateDeck() {
  const p = new PptxGenJS();
  p.layout = "LAYOUT_WIDE";
  const two = p.addSlide();
  two.background = { color: "F1F5F9" };
  two.addText("Title of the slide", { x: 0.5, y: 0.3, w: 12, h: 0.9, fontSize: 30, objectName: "Title" });
  two.addText("Left column", { x: 0.5, y: 1.5, w: 6, h: 4, fontSize: 16, objectName: "Left" });
  two.addText("Right column", { x: 6.8, y: 1.5, w: 6, h: 4, fontSize: 16, objectName: "Right" });
  two.addNotes("type: Two columns");
  const kpi = p.addSlide();
  kpi.addText("Key figure", { x: 0.5, y: 0.3, w: 12, h: 0.9, fontSize: 30, objectName: "Title" });
  kpi.addImage({ data: PIXEL, x: 1, y: 1.5, w: 11, h: 5, objectName: "Visual" });
  kpi.addNotes("type: Visual");
  return (await p.write({ outputType: "uint8array" })) as Uint8Array;
}

/** A workbook with one table "Tasks" (A1:C3) on Sheet1, as Excel writes them (minimal parts). */
export async function sampleWorkbook() {
  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  const x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${x}" xmlns:r="${r}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${r}/styles" Target="styles.xml"/></Relationships>`);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${x}"><fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs><tableStyles count="0"/></styleSheet>`);
  zip.file("xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${x}" xmlns:r="${r}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Title</t></is></c><c r="B1" t="inlineStr"><is><t>Status</t></is></c><c r="C1" t="inlineStr"><is><t>Due</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2" t="inlineStr"><is><t>Open</t></is></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`);
  zip.file("xl/worksheets/_rels/sheet1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/table" Target="../tables/table1.xml"/></Relationships>`);
  zip.file("xl/tables/table1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="${x}" id="1" name="Tasks" displayName="Tasks" ref="A1:C3"><autoFilter ref="A1:C3"/><tableColumns count="3"><tableColumn id="1" name="Title"/><tableColumn id="2" name="Status"/><tableColumn id="3" name="Due"/></tableColumns></table>`);
  return zip.generateAsync({ type: "uint8array" });
}
