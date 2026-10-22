/**
 * The building blocks of the slide templates: the deck's theme, the template contract (zod args, an
 * animated component, a native PowerPoint version) and small shared pieces (rich text, eyebrow, rule).
 * Slides are laid out in container units (cqw: % of the slide width), so they scale to any size.
 */
import type { ReactNode } from "react";
import { motion, type Variants } from "motion/react";
import type PptxGenJS from "pptxgenjs";
import { z } from "zod";

export const themeSchema = z.object({
  primary: z.string().default("#2563eb").describe("Accent color (hex)"),
  background: z.string().default("#ffffff").describe("Slide background (hex)"),
  text: z.string().default("#0f172a").describe("Text color (hex)"),
  fontTitle: z.string().default("Plus Jakarta Sans").describe("Font of the titles (a Google font)"),
  fontText: z.string().default("Plus Jakarta Sans").describe("Font of the text (a Google font)"),
});
export type Theme = z.infer<typeof themeSchema>;

/** What a template knows of the deck: its title, the slide's place, the section it is in. */
export type DeckInfo = { title: string; index: number; count: number; section: number; sections: string[] };

export type PptxContext = { theme: Theme; deck: DeckInfo; pptx: PptxGenJS };

export type Template<S extends z.ZodObject> = {
  id: string;
  /** For the AI and the user: what the slide is for. */
  description: string;
  args: S;
  /** Starts a part of the deck (numbered, listed in the outline). */
  section?: boolean;
  /** Full bleed: no header nor page number. */
  bare?: boolean;
  Component: (p: { args: z.infer<S>; theme: Theme; deck: DeckInfo }) => ReactNode;
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

/** The usual slide: eyebrow + title + rule on top, content under it. */
export function Frame({ eyebrow, title, children }: { eyebrow?: string; title: string; children: ReactNode }) {
  return (
    <M.div variants={stagger()} className="flex size-full flex-col gap-[2.2cqw] px-[6cqw] pt-[5cqw] pb-[5.5cqw]">
      <div className="flex flex-col gap-[1cqw]">
        {eyebrow && <Eyebrow>{eyebrow}</Eyebrow>}
        <Title>{title}</Title>
        <Rule />
      </div>
      <div className="min-h-0 flex-1">{children}</div>
    </M.div>
  );
}

/* --------------------------------- PowerPoint -------------------------------- */

/** Slide size in inches (16:9) and the usual title block. */
export const W = 13.333;
export const H = 7.5;
export function pptxTitle(slide: PptxGenJS.Slide, title: string, ctx: PptxContext, eyebrow?: string) {
  const { theme } = ctx;
  if (eyebrow) slide.addText(eyebrow.toUpperCase(), { x: 0.8, y: 0.55, w: 11, h: 0.35, fontFace: theme.fontText, fontSize: 11, bold: true, color: hex(theme.primary), charSpacing: 3 });
  slide.addText(plain(title), { x: 0.8, y: 0.9, w: 11.7, h: 0.9, fontFace: theme.fontTitle, fontSize: 32, bold: true, color: hex(theme.text) });
  slide.addShape("rect", { x: 0.8, y: 1.85, w: 0.8, h: 0.04, fill: { color: hex(theme.primary) }, line: { color: hex(theme.primary) } });
}
