import { analysisToExcel } from "@server/excel-export";
import readXlsxFile from "read-excel-file/node";
const a = JSON.parse(await Bun.file(".local/last-doc.json").text());
const buf = await analysisToExcel(a);
await Bun.write(".local/analysis.xlsx", buf);
const sheets = await readXlsxFile(buf);
for (const s of sheets) console.log(s.sheet, s.data.length, "rows:", JSON.stringify(s.data.slice(0, 3)).slice(0, 200));
