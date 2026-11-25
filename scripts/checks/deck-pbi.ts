import { deckToPptx } from "@server/pptx/deck";
import { slidePictures } from "@server/pptx/render";
import { closeBrowser } from "@server/capture/browser";
const md = `---
title: Revue Power BI
footer: Revue Power BI
---

# Revue du pipeline

Novembre 2026

---
layout: agenda
---

# Sommaire

1. Vue d'ensemble
2. Régions

---
layout: report
report: Revenue Opportunities
page: Opportunity Overview
---

# La page du rapport, en direct

---
layout: visual
report: Revenue Opportunities
visual: Revenue by Region and State
---

# East porte 43 % du pipeline

- **840 M$** pour East
- West reste en retrait
- À sécuriser en priorité

---
layout: compare
report: Revenue Opportunities
visual: Pipeline by Stage
filters: 'Account'[Region] = East
caption: East
filters2: 'Account'[Region] = West
caption2: West
---

# East et West : même entonnoir ?

---
layout: tutorial
report: Revenue Opportunities
visual: Revenue and Opportunity Count by Month
---

# Lire le graphique mensuel

1. Les barres : le revenu
2. La ligne : le nombre d'opportunités
3. Survolez un mois pour le détail

---
layout: end
---

Merci
`;
const t = Date.now();
const pptx = await deckToPptx(md);
console.log("pptx", pptx.length, Date.now() - t, "ms");
const pics = await slidePictures(new Uint8Array(pptx));
for (const [i, p] of pics.entries()) await Bun.write(`.local/p-${i + 1}.png`, p);
console.log(pics.length, Date.now() - t, "ms");
await closeBrowser();
