import { deckToPptx } from "@server/pptx/deck";
import { saveFile } from "@server/files";
const md = JSON.parse(await Bun.file(".local/deck-doc.json").text());
const f = await saveFile(new Uint8Array(await deckToPptx(md)), "revue-pipeline.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
console.log(JSON.stringify({ mode: "edit", doc: { fileId: f.id, name: "revue-pipeline.pptx" } }));
