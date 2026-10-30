/**
 * Example args for each template (Templates tab of the Presenter): written from its schema with lorem
 * ipsum and Unsplash photos, so a new template shows without an example of its own.
 */
import { z } from "zod";
import type { Template } from "./kit";

type Shape = { type?: string | string[]; properties?: Record<string, Shape>; required?: string[]; items?: Shape | Shape[]; prefixItems?: Shape[]; enum?: unknown[]; anyOf?: Shape[]; minItems?: number; maxItems?: number; minimum?: number; maximum?: number };

const PHOTOS = ["1497366216548-37526070297c", "1551288049-bebda4e38f71", "1522071820081-009f0129c71c", "1460925895917-afdab827c52f", "1504384308090-c894fdcc538d", "1519389950473-47ba0277781c"];
export const unsplash = (n: number) => `https://images.unsplash.com/photo-${PHOTOS[n % PHOTOS.length]}?w=1600&q=80&auto=format&fit=crop`;
const WORDS = "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua".split(" ");
const lorem = (n: number, from = 0) => {
  const text = Array.from({ length: n }, (_, i) => WORDS[(from + i) % WORDS.length]).join(" ");
  return text.charAt(0).toUpperCase() + text.slice(1);
};
const ICONS = ["rocket", "chart-bar", "users-round", "shield-check", "database", "sparkles"];

/** Optional args left out of the examples (colors per cell, Power BI zones): the plain version reads better. */
const SKIP = /Colors$|^crop$|^visuals?$|^page$|^current$|^headerColor$/;

function value(shape: Shape, key: string, n: number, report: string): unknown {
  if (shape.anyOf) return value(shape.anyOf.find((s) => s.type !== "null") ?? shape.anyOf[0], key, n, report);
  if (shape.enum) return shape.enum[0];
  const type = Array.isArray(shape.type) ? shape.type[0] : shape.type;
  if (type === "array") {
    const count = Math.min(shape.maxItems ?? 3, Math.max(shape.minItems ?? 0, key === "rows" ? 4 : 3));
    const item = Array.isArray(shape.items) ? shape.items[0] : (shape.items ?? shape.prefixItems?.[0] ?? { type: "string" });
    if (key === "rows") return Array.from({ length: count }, (_, i) => ["Lorem", `${12 + i * 3} k€`, `${(8 - i * 2).toFixed(0)} %`]);
    if (key === "columns" && (item.type === "string" || !item.type)) return ["Name", "Value", "Change"];
    return Array.from({ length: count }, (_, i) => value(item, key.replace(/s$/, ""), n + i, report));
  }
  if (type === "object") {
    const props = shape.properties ?? {};
    const out: Record<string, unknown> = {};
    for (const [k, s] of Object.entries(props)) {
      const required = shape.required?.includes(k);
      if (!required && SKIP.test(k)) continue;
      // One picture per card: an icon rather than an image or the mascot too.
      if (!required && (k === "image" || k === "mascot") && props.icon) continue;
      out[k] = value(s, k, n, report);
    }
    return out;
  }
  if (type === "number" || type === "integer") return Math.max(shape.minimum ?? 0, Math.min(shape.maximum ?? 9, n + 1));
  if (type === "boolean") return false;
  switch (key) {
    case "image":
      return unsplash(n);
    case "icon":
      return ICONS[n % ICONS.length];
    case "color":
      return "#2563EB";
    case "report":
      return report;
    case "value":
      return `${(n + 1) * 12} %`;
    case "title":
    case "heading":
    case "name":
    case "label":
      return lorem(3, n * 3);
    case "eyebrow":
    case "role":
    case "note":
    case "caption":
    case "context":
    case "footer":
    case "author":
      return lorem(2, n * 2 + 5);
    default:
      return `${lorem(10, n * 4)}.`;
  }
}

/** A slide of the template with example args (validated: the optional ones dropped if they do not fit). */
export function exampleSlide(template: Template<z.ZodObject>, report = "Sales report") {
  if (template.example) return { template: template.id, ...template.example };
  const shape = z.toJSONSchema(template.args, { unrepresentable: "any" }) as Shape;
  const args = value(shape, "", 0, report) as Record<string, unknown>;
  if (template.args.safeParse(args).success) return { template: template.id, ...args };
  const required = Object.fromEntries(Object.entries(args).filter(([k]) => shape.required?.includes(k)));
  return { template: template.id, ...required };
}

/** The template's arguments as a table: name, required, type in short, description. */
export function argRows(template: Template<z.ZodObject>) {
  const shape = z.toJSONSchema(template.args, { unrepresentable: "any" }) as Shape & { properties?: Record<string, Shape & { description?: string }> };
  const typeOf = (s: Shape): string => {
    if (s.anyOf) return s.anyOf.map(typeOf).filter((t) => t !== "null").join(" | ");
    if (s.enum) return s.enum.map(String).join(" | ");
    const type = Array.isArray(s.type) ? s.type.join(" | ") : s.type;
    if (type === "array") return `${typeOf((Array.isArray(s.items) ? s.items[0] : s.items) ?? {})}[]`;
    if (type === "object") return `{ ${Object.keys(s.properties ?? {}).join(", ")} }`;
    return type ?? "any";
  };
  return Object.entries(shape.properties ?? {}).map(([name, s]) => ({ name, required: !!shape.required?.includes(name), type: typeOf(s), description: s.description ?? "" }));
}
