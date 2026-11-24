/** A test Excel file (targets by region and quarter) saved as a server file; prints its FileRef. */
import writeXlsxFile from "write-excel-file/node";
import { saveFile } from "@server/files";
const rows: any[][] = [[{ value: "Objectifs commerciaux 2015" }], [], ["Région", "Trimestre", "Objectif CA"].map((v) => ({ value: v }))];
const targets: Record<string, number[]> = { East: [190e6, 210e6, 200e6, 240e6], Central: [160e6, 170e6, 180e6, 210e6], West: [110e6, 100e6, 95e6, 120e6] };
for (const [r, q] of Object.entries(targets)) q.forEach((v, i) => rows.push([{ value: r }, { value: `T${i + 1}` }, { value: v, type: Number }]));
const buf = await writeXlsxFile(rows as never).toBuffer();
console.log(JSON.stringify(await saveFile(new Uint8Array(buf), "objectifs-2015.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")));
