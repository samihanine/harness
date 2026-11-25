/** One example per layout of the template: shown in the Template tab (picture + markdown to copy). */
import type { Layout } from "./deck";

export const LAYOUT_EXAMPLES: { layout: Layout; title: string; group: "Template" | "Content" | "Power BI"; markdown: string }[] = [
  { layout: "cover", title: "Cover", group: "Template", markdown: "layout: cover\n---\n# Quarterly review\n\nSales & operations\n\nNovember 2026" },
  { layout: "agenda", title: "Agenda", group: "Template", markdown: "layout: agenda\nactive: 2\n---\n# Agenda\n\n1. Context\n2. Results\n3. Next steps" },
  { layout: "section", title: "Section", group: "Template", markdown: "layout: section\n---\n# Results" },
  { layout: "end", title: "End", group: "Template", markdown: "layout: end\n---\n# Thank you\n\nQuestions & discussion" },
  { layout: "default", title: "Content", group: "Content", markdown: "# The title says the message\n\n- A key point, in one line\n- A second point\n  - with a detail\n\n> **To remember**\n> One sentence the audience keeps." },
  { layout: "kpis", title: "Key figures", group: "Content", markdown: "layout: kpis\n---\n# Strong year, driven by the East\n\n- **$1.97 bn** Revenue — +12 % vs 2025\n- **458** Opportunities\n- **43 %** East's share" },
  { layout: "two-cols", title: "Two columns", group: "Content", markdown: "layout: two-cols\n---\n# Two columns\n\nText on the left.\n\n- one\n- two\n\n::right::\n\n| Region | Revenue |\n|---|---|\n| East | 840 M |\n| West | 404 M |" },
  { layout: "columns", title: "Columns", group: "Content", markdown: "layout: columns\n---\n# Three next steps\n\n::col::\n### Qualify\n- Sort the leads\n\n::col::\n### Speed up\n- Secure the big deals\n\n::col::\n### Rebalance\n- Support the West" },
  { layout: "table", title: "Table with pictures", group: "Content", markdown: "layout: table\n---\n# Our flagship products\n\n| Product | Picture | Share |\n|---|---|---|\n| Bag | ![](https://upload.wikimedia.org/wikipedia/commons/a/a7/React-icon.svg) | 32 % |\n| Scarf | ![](https://upload.wikimedia.org/wikipedia/commons/a/a7/React-icon.svg) | 21 % |" },
  { layout: "default", title: "Chart", group: "Content", markdown: "# Revenue by month\n\n```chart\n{\"type\":\"column\",\"labels\":[\"Jan\",\"Feb\",\"Mar\",\"Apr\"],\"series\":[{\"name\":\"Revenue\",\"values\":[1.2,1.5,1.4,1.9]}]}\n```" },
  { layout: "image-right", title: "Picture beside text", group: "Content", markdown: "layout: image-right\nimage: https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/Tour_Eiffel_Wikimedia_Commons.jpg/640px-Tour_Eiffel_Wikimedia_Commons.jpg\n---\n# A picture and its message\n\n- The picture fills half of the slide\n- The text stays on the other half" },
  { layout: "full-image", title: "Full picture", group: "Content", markdown: "layout: full-image\nimage: https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/Tour_Eiffel_Wikimedia_Commons.jpg/640px-Tour_Eiffel_Wikimedia_Commons.jpg\n---\n# A title over the picture" },
  { layout: "quote", title: "Quote", group: "Content", markdown: "layout: quote\n---\n> Quality is remembered long after the price is forgotten.\n\n— A founder" },
  { layout: "report", title: "Report page", group: "Power BI", markdown: "layout: report\nreport: {report}\n---\n# The report, live" },
  { layout: "visual", title: "Visual + message", group: "Power BI", markdown: "layout: visual\nreport: {report}\nvisual: {visual}\n---\n# What the visual shows\n\n- The key figure\n- What it means\n- What to do" },
  { layout: "compare", title: "Compare", group: "Power BI", markdown: "layout: compare\nreport: {report}\nvisual: {visual}\nfilters: {filters}\ncaption: Left\nfilters2:\ncaption2: All\n---\n# Two views side by side" },
  { layout: "tutorial", title: "Tutorial", group: "Power BI", markdown: "layout: tutorial\nreport: {report}\nvisual: {visual}\n---\n# How to read this chart\n\n1. Choose the period in the slicer\n2. Read the trend\n3. Hover a bar for the detail" },
];
