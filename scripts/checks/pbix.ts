/**
 * The Builder's .pbix path: a file imported privately (My workspace), changed through its copy's definition,
 * then rebuilt with that definition (connection kept).  bun scripts/checks/pbix.ts <file.pbix> [--thin]
 * --thin first turns the file into a report live-connected to its model (no DataModel), as a Build-only user has.
 */
import JSZip from "jszip";
import { forgetPbixCopy, importPbix, pbixInfo, pbixWithDefinition } from "@server/pbi/pbix";
import { draftOf, pageOf, saveDraft, writeJson } from "@server/pbi/draft";
import { buildVisual, newName } from "@server/pbi/visual-spec";
let data: Uint8Array = new Uint8Array(await Bun.file(process.argv[2]).arrayBuffer());
if (process.argv.includes("--thin")) {
  const zip = await JSZip.loadAsync(data);
  const datasetId = JSON.parse((await zip.file("Connections")!.async("string")).replace(/^﻿/, "")).RemoteArtifacts[0].DatasetId;
  zip.remove("DataModel");
  zip.remove("DiagramLayout");
  zip.file("Connections", JSON.stringify({ Version: 3, Connections: [{ Name: "EntityDataSource", ConnectionString: `Data Source=pbiazure://api.powerbi.com;Initial Catalog=${datasetId};Identity Provider="https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc";Integrated Security=ClaimsToken`, ConnectionType: "pbiServiceLive", PbiServiceModelId: null, PbiModelVirtualServerName: "sobe_wowvirtualserver", PbiModelDatabaseName: datasetId }] }));
  const types = (await zip.file("[Content_Types].xml")!.async("string")).replace(/<Override PartName="\/(DataModel|DiagramLayout)"[^>]*\/>/g, "");
  zip.file("[Content_Types].xml", types);
  data = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  await Bun.write(".local/thin.pbix", data);
}
console.log("info:", await pbixInfo(data));
let t = Date.now();
const copy = await importPbix(data, `[Hub check] ${Date.now()}`);
console.log("imported in", ((Date.now() - t) / 1000).toFixed(1), "s:", copy);
t = Date.now();
const d = await draftOf({}, copy);
const page = pageOf(d, d.parts.find((p) => p.path.endsWith("/page.json"))!.path.split("/")[2]);
const name = newName();
writeJson(d, `definition/pages/${page.name}/visuals/${name}/visual.json`, buildVisual({ type: "textbox", x: 20, y: 640, width: 500, height: 60, text: "Changed by the Builder", textStyle: { bold: true, color: "#B39251", fontSize: 20 } }, name, 9000, undefined, []));
await saveDraft(d);
console.log("saved in", ((Date.now() - t) / 1000).toFixed(1), "s (legacy:", !!d.legacy, ")");
const out = await pbixWithDefinition(data, copy);
await Bun.write(".local/out.pbix", out);
const z = await JSZip.loadAsync(out);
console.log("rebuilt:", Object.keys(z.files).join(", ").slice(0, 400));
console.log("info of the rebuilt file:", await pbixInfo(out));
if (!process.argv.includes("--keep")) await forgetPbixCopy(copy.id);
