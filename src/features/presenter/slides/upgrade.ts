/**
 * Slides of older versions in today's templates: the "section" page is an agenda naming its part, the
 * "issues" page (issue → solution) is the causality page (cause → effect). Used when a deck is read
 * (store.ts) and checked (index.ts).
 */
export function upgradeSlide(s: unknown): unknown {
  const x = s as Record<string, unknown> | undefined;
  if (x?.template === "section") return { template: "agenda", part: x.title, ...(x.image ? { image: x.image } : {}) };
  if (x?.template === "issues") {
    const items = Array.isArray(x.items) ? (x.items as Record<string, unknown>[]).map((i) => ({ cause: i?.issue ?? i?.cause, effect: i?.solution ?? i?.effect })) : x.items;
    return { ...x, template: "causality", items };
  }
  return s;
}
