/**
 * The building blocks of the slide templates: the deck's theme, the template contract (zod args, an
 * animated component, a native PowerPoint version) and small shared pieces (rich text, eyebrow, rule).
 * Slides are laid out in container units (cqw: % of the slide width), so they scale to any size.
 */
import { createContext, useContext, type ReactNode } from "react";
import { Markdown } from "@/components/markdown";
import { BRAND_NAME, PRIMARY_COLOR } from "@/config";
import { motion, type Variants } from "motion/react";
import type PptxGenJS from "pptxgenjs";
import { z } from "zod";

/** Header variants of the content pages (deck's general config). */
export const HEADERS = ["none", "breadcrumb", "rule", "brand"] as const;
/** Variants of older versions, by the one now closest. */
const OLD_HEADERS: Record<string, (typeof HEADERS)[number]> = { progress: "rule", monogram: "breadcrumb" };

/** The app's primary color (Settings) when it is a hex color: the default accent of a deck. */
export const appPrimary = () => (/^#[0-9a-f]{6}$/i.test(PRIMARY_COLOR.trim()) ? PRIMARY_COLOR.trim() : "#2563eb");

export const themeSchema = z.object({
  primary: z.string().default(appPrimary).describe("Accent color (hex)"),
  background: z.string().default("#ffffff").describe("Slide background (hex)"),
  text: z.string().default("#0f172a").describe("Text color (hex)"),
  fontTitle: z.string().default("Plus Jakarta Sans").describe("Font of the titles (a Google font)"),
  fontText: z.string().default("Plus Jakarta Sans").describe("Font of the text (a Google font)"),
  secondary: z.string().default("#8b5a2b").describe("Second accent (numbered badges, cards, bands), hex"),
  success: z.string().default("#15803d").describe("Color of tips / info items, hex"),
  header: z
    .preprocess((v) => (typeof v === "string" ? (OLD_HEADERS[v] ?? v) : v), z.enum(HEADERS))
    .default("breadcrumb")
    .describe("Header on top of every content page: the brand letter, the page title on the left, then none (no header: the title is on the page), breadcrumb (deck › part › subpart on the right), rule (the same over a fine line), brand (the title with its number in the subpart, e.g. \"Sales page (2/6)\", and the brand name on the right, over a fine line)"),
  brand: z.string().optional().describe("Brand name of the deck (headers, covers, monogram letter); default: the app's brand (Settings)"),
});
export type Theme = z.infer<typeof themeSchema>;

/**
 * The kind of a page: title pages (cover, section, subsection, end), the summary, or a content page.
 * The outline comes from the section and subsection pages; content pages get the header.
 */
export type Kind = "cover" | "section" | "subsection" | "summary" | "content" | "end";

/** What a template knows of the deck: its title, the slide's place, the part and subpart it is in, the outline. */
export type DeckInfo = {
  title: string;
  index: number;
  count: number;
  section: number;
  sections: string[];
  /** Subpart in the current part (-1: none yet) and the subparts of that part. */
  subsection: number;
  subsections: string[];
  /** The whole outline: each part with its subparts (summary pages). */
  outline: { title: string; subsections: string[] }[];
  /** The part a summary placed here leads to (0 before the first part). */
  next: number;
  /** This content page's place among the content pages of its subpart (or part): "(2/6)" of the brand header. */
  group: { index: number; count: number };
};

/** The deck's brand: its theme's, else the app's default (config.ts, Settings › General). */
export const brandOf = (theme: { brand?: string }) => theme.brand?.trim() || BRAND_NAME;

/** Set by the stage: the "brand" header shows the page title, so content pages leave it out. */
export const SlideContext = createContext({ titleInHeader: false });

/**
 * A text size (cqw) that keeps texts inside their box: from the box's area (cqw²) and the amount of
 * text, between `min` and `max` (a long list gets smaller text instead of overflowing).
 */
export function fitText(texts: string[], area: number, { min = 0.9, max = 1.6, lineGap = 1.6 } = {}) {
  const chars = texts.reduce((n, t) => n + plain(t).length, 0) + texts.length * 18;
  return Math.max(min, Math.min(max, Math.sqrt(area / (chars * 0.55 * lineGap))));
}

/** Markdown inside a slide (paragraphs, lists, **bold**), sized by its parent. */
export const SlideMarkdown = ({ text }: { text: string }) => (
  <div className="slide-markdown [&_.markdown]:text-[1em] [&_li]:mt-[0.2em] [&_p]:my-0 [&_strong]:font-semibold">
    <Markdown>{text}</Markdown>
  </div>
);

/** `toile`: the title pages' flowers as a PNG (data), drawn in the deck's colors before the export. */
export type PptxContext = { theme: Theme; deck: DeckInfo; pptx: PptxGenJS; toile?: string };

export type Template<S extends z.ZodObject> = {
  id: string;
  /** For the AI and the user: what the slide is for. */
  description: string;
  args: S;
  /** Default: a content page. */
  kind?: Kind;
  /** The kind given by the slide's args (e.g. an agenda naming a part starts that part). */
  kindOf?: (args: Record<string, unknown>) => Kind;
  /** Full bleed: no header nor page number. */
  bare?: boolean;
  Component: (p: { args: z.infer<S>; theme: Theme; deck: DeckInfo }) => ReactNode;
  /** Example args (Settings › Slides) when the ones written from the schema would not say much. */
  example?: Record<string, unknown>;
  pptx: (slide: PptxGenJS.Slide, args: z.infer<S>, ctx: PptxContext) => void;
};

export const defineTemplate = <S extends z.ZodObject>(t: Template<S>) => t;

/* ---------------------------------- Motion ---------------------------------- */

export const ease = [0.22, 1, 0.36, 1] as const;
/** Children appear one after the other. */
export const stagger = (delay = 0.08, start = 0.15): Variants => ({ hidden: {}, show: { transition: { staggerChildren: delay, delayChildren: start } } });
export const rise: Variants = { hidden: { opacity: 0, y: "1.2cqw" }, show: { opacity: 1, y: 0, transition: { duration: 0.7, ease } } };
export const fade: Variants = { hidden: { opacity: 0 }, show: { opacity: 1, transition: { duration: 0.8, ease } } };
export const reveal: Variants = { hidden: { clipPath: "inset(0 0 0 100%)" }, show: { clipPath: "inset(0 0 0 0%)", transition: { duration: 1.2, ease } } };

export const M = motion;

/* ----------------------------------- Text ----------------------------------- */

/** **bold** and ==highlight== (accent color) in a line of text. */
export function Rich({ text }: { text: string }) {
  const parts = text.split(/(\*\*[^*]+\*\*|==[^=]+==)/g);
  return (
    <>
      {parts.map((p, i) =>
        p.startsWith("**") ? (
          <strong key={i} className="font-semibold">{p.slice(2, -2)}</strong>
        ) : p.startsWith("==") ? (
          <span key={i} className="font-semibold" style={{ color: "var(--s-primary)" }}>{p.slice(2, -2)}</span>
        ) : (
          p
        ),
      )}
    </>
  );
}

/** The same text without its marks (PowerPoint gets runs instead). */
export const plain = (text: string) => text.replace(/\*\*([^*]+)\*\*/g, "$1").replace(/==([^=]+)==/g, "$1");

/** Text runs for PowerPoint: bold and highlighted parts. */
export function runs(text: string, theme: Theme): PptxGenJS.TextProps[] {
  return text
    .split(/(\*\*[^*]+\*\*|==[^=]+==)/g)
    .filter(Boolean)
    .map((p) => (p.startsWith("**") ? { text: p.slice(2, -2), options: { bold: true } } : p.startsWith("==") ? { text: p.slice(2, -2), options: { bold: true, color: hex(theme.primary) } } : { text: p }));
}

export const hex = (color: string) => color.replace("#", "").slice(0, 6).toUpperCase();
/** A color between two (`p`: the share of the first), as PowerPoint wants it: the accent's tints and shades. */
export function mix(a: string, b: string, p: number) {
  const rgb = (c: string) => [0, 2, 4].map((i) => parseInt(hex(c).slice(i, i + 2), 16) || 0);
  const [x, y] = [rgb(a), rgb(b)];
  return x.map((v, i) => Math.round(v * p + y[i] * (1 - p)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/* ---------------------------- Tints (on screen) ---------------------------- */

/** The accent mixed with the background (`p` %): fills of cards, bands. */
export const tint = (p: number) => `color-mix(in srgb, var(--s-primary) ${p}%, var(--s-bg))`;
/** The accent darkened by the text color: readable on a light background (a pale accent is not). */
export const deep = "color-mix(in srgb, var(--s-primary) 62%, var(--s-text))";
/** The text color faded (`p` %): lines, muted texts. */
export const faint = (p: number) => `color-mix(in srgb, var(--s-text) ${p}%, transparent)`;
/** The content box of a page under its title (cqw): templates drawn with positions fit it. */
export const STAGE = { w: 88, h: 34 };
export const pad = (n: number) => String(n).padStart(2, "0");

export function Eyebrow({ children }: { children: ReactNode }) {
  return (
    <M.div variants={rise} className="text-[1.05cqw] font-semibold tracking-[0.18em] uppercase" style={{ color: "var(--s-primary)" }}>
      {children}
    </M.div>
  );
}

export function Rule({ className = "" }: { className?: string }) {
  return <M.div variants={{ hidden: { scaleX: 0 }, show: { scaleX: 1, transition: { duration: 0.9, ease } } }} className={`h-[0.22cqw] w-[6cqw] origin-left rounded-full ${className}`} style={{ background: "var(--s-primary)" }} />;
}

export function Title({ children, size = 3.6 }: { children: ReactNode; size?: number }) {
  return (
    <M.h2 variants={rise} className="leading-[1.08] font-bold tracking-tight" style={{ fontSize: `${size}cqw`, fontFamily: "var(--s-title)" }}>
      {children}
    </M.h2>
  );
}

/** The usual slide: title + rule on top, content under it (a header shows the title itself, on its left). */
export function Frame({ title, aside, children }: { title: string; aside?: ReactNode; children: ReactNode }) {
  const { titleInHeader } = useContext(SlideContext);
  if (titleInHeader)
    return (
      <M.div variants={stagger()} className="flex size-full flex-col px-[6cqw] pt-[8cqw] pb-[4.5cqw]">
        <div className="min-h-0 flex-1">{children}</div>
      </M.div>
    );
  return (
    <M.div variants={stagger()} className="flex size-full flex-col gap-[2.2cqw] px-[6cqw] pt-[5.8cqw] pb-[5cqw]">
      <div className="flex items-end gap-[2cqw]">
        <div className="flex min-w-0 flex-1 flex-col gap-[1cqw]">
          <Title>{title}</Title>
          <Rule />
        </div>
        {aside}
      </div>
      <div className="min-h-0 flex-1">{children}</div>
    </M.div>
  );
}

/* --------------------------------- PowerPoint -------------------------------- */

type Point = { x: number; y: number };
/**
 * A line, polyline or curve from points placed on the slide (inches): custom geometry in their box.
 * `c1`, `c2`: the control points of a cubic curve ending at the point.
 */
export function pptxPath(slide: PptxGenJS.Slide, points: (Point & { c1?: Point; c2?: Point })[], options: { color: string; width?: number; dash?: "dash" | "sysDot"; arrow?: boolean; fill?: string; transparency?: number }) {
  const all = points.flatMap((p) => [p, ...(p.c1 ? [p.c1, p.c2!] : [])]);
  const x = Math.min(...all.map((p) => p.x));
  const y = Math.min(...all.map((p) => p.y));
  const w = Math.max(0.01, Math.max(...all.map((p) => p.x)) - x);
  const h = Math.max(0.01, Math.max(...all.map((p) => p.y)) - y);
  const at = (p: Point) => ({ x: p.x - x, y: p.y - y });
  slide.addShape("custGeom" as PptxGenJS.ShapeType, {
    x,
    y,
    w,
    h,
    points: [...points.map((p, i) => (p.c1 && i ? { ...at(p), curve: { type: "cubic" as const, x1: at(p.c1).x, y1: at(p.c1).y, x2: at(p.c2!).x, y2: at(p.c2!).y } } : { ...at(p), ...(i ? {} : { moveTo: true }) })), ...(options.fill ? [{ close: true as const }] : [])],
    line: { color: options.color, width: options.width ?? 1, dashType: options.dash, endArrowType: options.arrow ? "triangle" : undefined, transparency: options.transparency },
    fill: options.fill ? { color: options.fill } : { type: "none" } as never,
  });
}

/** Slide size in inches (16:9) and the usual title block. */
export const W = 13.333;
export const H = 7.5;
export function pptxTitle(slide: PptxGenJS.Slide, title: string, ctx: PptxContext) {
  const { theme } = ctx;
  // With a header, the title is in it (pptx.ts).
  if (theme.header !== "none") return;
  slide.addText(plain(title), { x: 0.8, y: 0.55, w: 11.7, h: 0.75, fontFace: theme.fontTitle, fontSize: 32, bold: true, color: hex(theme.text), valign: "middle" });
  slide.addShape("rect", { x: 0.8, y: 1.33, w: 0.8, h: 0.04, fill: { color: hex(theme.primary) }, line: { color: hex(theme.primary) } });
}
