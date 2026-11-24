import { newDeck } from "@shared/deck";
import { deckToPptx } from "@server/pptx/deck";
import { slidePictures } from "@server/pptx/render";
import { closeBrowser } from "@server/capture/browser";
const md = newDeck("Revenue review") + `
---
layout: kpis
---

# Key figures

- **$1.97 bn** Opportunity revenue — +12 % vs 2025
- **444** Opportunities
- **East** First region — 43 % of revenue

---
layout: two-cols
---

# Two columns

Left text with **bold**.

- one
- two

::right::

| Region | Revenue | Logo |
|---|---|---|
| East | 0.84 bn | ![](https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/React-icon.svg/64px-React-icon.svg.png) |
| West | 0.40 bn | |

---

# Chart

\`\`\`chart
{"type":"column","labels":["Jan","Feb","Mar"],"series":[{"name":"Revenue","values":[1,2,3]}]}
\`\`\`

---
layout: end
---

Thank you
`;
const t = Date.now();
const pptx = await deckToPptx(md);
console.log("pptx", pptx.length, Date.now() - t);
await Bun.write(".local/check.pptx", pptx);
const pics = await slidePictures(new Uint8Array(pptx));
console.log("pictures", pics.length, Date.now() - t);
for (const [i, p] of pics.entries()) await Bun.write(`.local/check-${i + 1}.png`, p);
await closeBrowser();
