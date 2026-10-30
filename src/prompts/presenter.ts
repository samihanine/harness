/** The Presenter: its instructions and the prompt copied for an outside assistant. */
import { templatesText } from "@/features/presenter/slides";
import { DAX_RULES } from "./dax";

/** The Presenter, with what the deck is written from. */
export const slidesInstructions = (chosen: { reports: unknown[]; datasets: unknown[] }) =>
  `Screen: SLIDES. You write a presentation as a short JSON: a theme and slides, each slide = a template id and its args (the templates draw it; you only give the content). Slides are numbered from 1.
Templates:
${templatesText()}
Structure: a cover page (cover — a picture and the brand —, title — engraved flowers — or hero), a summary page (filled automatically from the parts), then each part: an agenda page naming it ("part": it starts the part and shows the progress), optional subtitle pages for its subparts (each shows the trail of the part), and its content pages; a closing page (end, or title for "Questions?"). The outline and the header of content pages come from these title pages: never write the breadcrumb yourself. Header style, colors (primary, secondary for numbered badges and cards, success for tips), fonts, brand name: set_deck (theme.header: none | breadcrumb | rule | brand — every header but none shows the page title on its left, so content pages carry their title there; brand adds its number in the subpart and the brand name). A report page explained step by step: tutorial (each item frames its own visual or group of visuals by their ids, numbered on the page: one item per visual, never two items on the same visual).
Rules: one idea per slide, short texts (a title of ≤ 8 words, items of ≤ 15 words); **bold** and ==highlight== (accent color) in texts; numbers in a grid → the table template.${
      chosen.reports.length ? " Power BI slides (report: a whole page; visual: one visual by its crop; compare: two): the report id, the page name and the crop come from the context, exactly." : " No report is chosen: no report / visual / compare slide (the user can choose reports in the Reports tab)."
    }${chosen.datasets.length ? ` Figures: compute them with run_dax on the chosen models (exact values), never invent them.\n${DAX_RULES}` : ""}`;

/** "Copy the prompt": for an outside assistant writing the deck's JSON. */
export const deckPrompt = (templates: string, current: string) =>
  `You write a slide deck as JSON for a presentation app. I will tell you what I want after this message.

Return ONLY the complete JSON of the deck (no explanation, no markdown fence). Write every text of the slides in the language of my request.

Format: { "title": string, "theme": { "primary", "secondary", "success", "background", "text" (hex colors), "fontTitle", "fontText" (Google font names), "header": "none" | "breadcrumb" | "rule" | "brand", "brand"? }, "slides": [ { "template": "<id>", ...its arguments } ] }

Rules:
- Structure: a cover (cover, title or hero), a summary or agenda, then for each part an agenda page naming it ("part"), optional subtitle pages and its content pages, and a closing page (end).
- One idea per slide; short texts (titles of 8 words at most, items of 15 words at most). In texts: **bold** and ==highlight==.
- Images: a public https link (e.g. https://images.unsplash.com/…) or an inline "data:image/png;base64,…".
- Power BI slides (report, visual, tutorial, compare) only with the report ids already in the current JSON.
- Arguments marked "?" are optional; use only the arguments listed.

## Templates

${templates}

## Current deck

${current}
`;
