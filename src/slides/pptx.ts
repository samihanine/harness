/** The deck as a native, editable PowerPoint file (shapes and text; live reports become links). */
import PptxGenJS from "pptxgenjs";
import { hex } from "./kit";
import { checkSlide, deckInfo, type Deck } from "./index";
import { loadReportLinks } from "./visual";

export async function downloadPptx(deck: Deck, name: string) {
  await loadReportLinks();
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = deck.title;
  const infos = deckInfo(deck);
  deck.slides.forEach((s, i) => {
    const { template, args } = checkSlide(s);
    const slide = pptx.addSlide();
    slide.background = { color: hex(deck.theme.background) };
    template.pptx(slide, args, { theme: deck.theme, deck: infos[i], pptx });
    // The header of content pages: where the page is (deck › part › subpart).
    const info = infos[i];
    if (!template.bare && (template.kind ?? "content") === "content" && deck.theme.header !== "none") {
      const trail = [deck.title, info.sections[info.section], info.subsection >= 0 ? info.subsections[info.subsection] : undefined].filter(Boolean).join("  ›  ");
      slide.addText(trail.toUpperCase(), { x: 0.8, y: 0.12, w: 11.7, h: 0.3, fontSize: 8, charSpacing: 2, color: hex(deck.theme.primary), fontFace: deck.theme.fontText });
    }
    if (!template.bare) slide.addText(`${i + 1}`, { x: 12.3, y: 6.9, w: 0.6, h: 0.3, fontSize: 9, align: "right", color: hex(deck.theme.text), transparency: 50, fontFace: deck.theme.fontText });
  });
  await pptx.writeFile({ fileName: `${name.replace(/[^\p{L}\p{N} _-]+/gu, "-")}.pptx` });
}
