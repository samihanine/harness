/**
 * The Template tab: one slide per layout, drawn as PowerPoint draws them, with the markdown to copy. The
 * Power BI layouts use a report the user can open (its first page's first visual). Drawn once, kept an hour.
 */
import { LAYOUT_EXAMPLES } from "@shared/deck-examples";
import { saveFile } from "../files";
import { reportCatalog } from "../pbi/catalog";
import { reportStructure } from "../pbi/report-structure";
import { deckToPptx } from "./deck";
import { slidePictures } from "./render";

let cache: { at: number; value: Promise<{ examples: { layout: string; title: string; group: string; markdown: string; picture?: string }[] }> } | undefined;

async function examples() {
  // A real report for the Power BI examples.
  let report = "", visual = "", filters = "";
  const first = (await reportCatalog().catch(() => []))[0];
  if (first) {
    report = first.name;
    const s = (await reportStructure({ id: first.id, groupId: first.groupId, appId: first.appId }).catch(() => undefined)) as { pages?: { visuals: { title?: string; type: string }[] }[] } | undefined;
    visual = s?.pages?.[0]?.visuals.find((v) => v.title && !/slicer|textbox|image|shape|button/i.test(v.type))?.title ?? "";
  }
  return LAYOUT_EXAMPLES.map((e) => ({ ...e, markdown: e.markdown.replace("{report}", report || "<report name>").replace("{visual}", visual || "<visual title>").replace("{filters}", filters) }));
}

export function templateGallery() {
  if (!cache || Date.now() - cache.at > 3600_000) {
    const value = (async () => {
      const list = await examples();
      const md = `---\ntitle: Template\nfooter: Template\n---\n${list.map((e) => e.markdown).join("\n\n---\n")}\n`;
      const pics = await slidePictures(new Uint8Array(await deckToPptx(md)), undefined, 960).catch(() => [] as Buffer[]);
      const files = await Promise.all(pics.map((p, i) => saveFile(p, `template-${i + 1}.png`, "image/png")));
      return { examples: list.map((e, i) => ({ ...e, picture: files[i]?.id })) };
    })();
    cache = { at: Date.now(), value };
    value.catch(() => (cache = undefined));
  }
  return cache.value;
}
