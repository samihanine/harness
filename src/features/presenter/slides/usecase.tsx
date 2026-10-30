import { CheckIcon, QuoteIcon } from "lucide-react";
import { z } from "zod";
import { Mascot } from "@/features/mascot";
import { pptxImage } from "./images";
import { Icon, outfitSchema, pptxPicture } from "./picture";
import { Frame, M, Rich, deep, defineTemplate, ease, fade, faint, hex, mix, plain, pptxTitle, rise, stagger, tint, type PptxContext } from "./kit";
import type PptxGenJS from "pptxgenjs";

/** A use case's args (also those of the conversation): who asks, what the assistant does for them. */
export const usecaseArgs = z.object({
  title: z.string(),
  mascot: outfitSchema.default("diver"),
  description: z.string().optional().describe("One line: what the assistant does for them"),
  personas: z.array(z.object({ role: z.string(), quote: z.string().describe("A typical question they ask"), icon: z.string().optional().describe("lucide icon name, e.g. briefcase") })).min(1).max(3),
  features: z.array(z.string()).min(1).max(5),
});

export const USECASE_EXAMPLE = {
  title: "Comprendre un rapport",
  mascot: "trainer",
  description: "Le Trainer explique chaque visuel et répond à partir des **données du modèle**.",
  personas: [
    { role: "Directeur commercial", quote: "Pourquoi les ventes Europe baissent ce mois-ci ?", icon: "briefcase" },
    { role: "Contrôleur de gestion", quote: "D'où vient ce chiffre de marge ?", icon: "calculator" },
    { role: "Nouvel arrivant", quote: "Comment lire cette page ?", icon: "graduation-cap" },
  ],
  features: ["Explique un visuel en un clic", "Répond à partir des données", "Applique les filtres demandés", "Guide pas à pas chaque page"],
};

/** A persona's round avatar: its icon on the accent. */
export function Avatar({ icon, size = 4.6 }: { icon?: string; size?: number }) {
  return (
    <span className="grid shrink-0 place-items-center rounded-full text-white shadow-[0_0.3cqw_0.8cqw_rgba(0,0,0,0.15)]" style={{ width: `${size}cqw`, height: `${size}cqw`, background: `linear-gradient(140deg, var(--s-primary), ${deep})` }}>
      <Icon name={icon ?? "user-round"} style={{ width: `${size * 0.46}cqw`, height: `${size * 0.46}cqw` }} />
    </span>
  );
}

/** A persona's avatar in PowerPoint: the accent circle and its icon in white. */
export function pptxAvatar(slide: PptxGenJS.Slide, icon: string | undefined, box: { x: number; y: number; d: number }, ctx: PptxContext) {
  slide.addShape("ellipse", { x: box.x, y: box.y, w: box.d, h: box.d, fill: { color: hex(ctx.theme.primary) }, line: { type: "none" } as never });
  const picture = pptxImage(`icon-white:${icon?.trim() || "user-round"}`, { x: box.x + box.d * 0.25, y: box.y + box.d * 0.25, w: box.d * 0.5, h: box.d * 0.5 });
  if (picture) slide.addImage(picture);
}

/** A use case: the personas as large cards (avatar, role, their question), what the assistant does as a band under them, the mascot on the right. */
export const usecase = defineTemplate({
  id: "usecase",
  description: "A use case of the app, personas first: 1-3 persona cards (role, a typical question they ask, a lucide icon), an optional line on what the assistant does, 1-5 features as a band under the cards, the mascot (outfit) on the right.",
  args: usecaseArgs,
  example: USECASE_EXAMPLE,
  Component: ({ args }) => (
    <Frame title={args.title}>
      <div className="grid h-full grid-cols-[78fr_22fr] gap-[2cqw]">
        <M.div variants={stagger(0.14, 0.2)} className="flex min-h-0 flex-col justify-center gap-[1.8cqw]">
          {args.description && (
            <M.p variants={rise} className="text-[1.45cqw] leading-snug opacity-75">
              <Rich text={args.description} />
            </M.p>
          )}
          <div className="grid min-h-0 gap-[1.6cqw]" style={{ gridTemplateColumns: `repeat(${args.personas.length}, minmax(0, 1fr))` }}>
            {args.personas.map((p, i) => (
              <M.div key={i} variants={rise} whileHover={{ y: "-0.6cqw" }} className="relative flex flex-col gap-[1.2cqw] overflow-hidden rounded-[1.2cqw] px-[1.8cqw] pt-[1.8cqw] pb-[1.6cqw] shadow-[0_0.4cqw_1.4cqw_rgba(0,0,0,0.07)]" style={{ background: "var(--s-bg)", border: `1px solid ${faint(10)}` }}>
                <span className="absolute inset-x-0 top-0 h-[0.35cqw]" style={{ background: "var(--s-primary)" }} />
                <div className="flex items-center gap-[1cqw]">
                  <Avatar icon={p.icon} />
                  <span className="text-[1.35cqw] leading-tight font-bold">{p.role}</span>
                </div>
                <QuoteIcon className="shrink-0" style={{ width: "2.4cqw", height: "2.4cqw", color: "var(--s-primary)" }} fill="currentColor" strokeWidth={0} />
                <p className="text-[1.5cqw] leading-snug italic" style={{ fontFamily: "var(--s-title)" }}>
                  <Rich text={p.quote} />
                </p>
              </M.div>
            ))}
          </div>
          <M.div variants={rise} className="flex flex-wrap items-center gap-[0.8cqw] rounded-[1cqw] px-[1.4cqw] py-[1.1cqw]" style={{ background: tint(16) }}>
            <span className="mr-[0.4cqw] text-[0.95cqw] font-semibold tracking-[0.14em] uppercase" style={{ color: deep }}>
              L'assistant
            </span>
            {args.features.map((f, i) => (
              <M.span key={i} initial={{ opacity: 0, scale: 0.85 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.9 + i * 0.12, duration: 0.5, ease }} className="flex items-center gap-[0.5cqw] rounded-full px-[1cqw] py-[0.45cqw] text-[1.15cqw]" style={{ background: "var(--s-bg)" }}>
                <CheckIcon style={{ width: "1.2cqw", height: "1.2cqw", color: deep }} strokeWidth={3} />
                <Rich text={f} />
              </M.span>
            ))}
          </M.div>
        </M.div>
        <M.div variants={fade} className="flex min-h-0 items-end justify-center">
          <Mascot outfit={args.mascot} state="talking" size={400} className="h-auto max-h-full w-full" />
        </M.div>
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const text = hex(theme.text);
    if (args.description) slide.addText(plain(args.description), { x: 0.8, y: 2.0, w: 9.1, h: 0.45, fontSize: 14, color: text, fontFace: theme.fontText, transparency: 20 });
    const gap = 0.25;
    const w = (9.1 - gap * (args.personas.length - 1)) / args.personas.length;
    args.personas.forEach((p, i) => {
      const x = 0.8 + i * (w + gap);
      const y = 2.6;
      slide.addShape("roundRect", { x, y, w, h: 2.7, fill: { color: hex(theme.background) }, line: { color: mix(theme.text, theme.background, 0.12), width: 0.75 }, rectRadius: 0.1 });
      slide.addShape("rect", { x: x + 0.05, y, w: w - 0.1, h: 0.05, fill: { color: hex(theme.primary) }, line: { type: "none" } as never });
      pptxAvatar(slide, p.icon, { x: x + 0.25, y: y + 0.3, d: 0.55 }, ctx);
      slide.addText(p.role, { x: x + 0.9, y: y + 0.3, w: w - 1.05, h: 0.55, fontSize: 13, bold: true, color: text, fontFace: theme.fontText, valign: "middle" });
      slide.addText("\u201C", { x: x + 0.2, y: y + 0.9, w: 0.6, h: 0.5, fontSize: 40, bold: true, color: hex(theme.primary), fontFace: theme.fontTitle });
      slide.addText(plain(p.quote), { x: x + 0.25, y: y + 1.4, w: w - 0.5, h: 1.2, fontSize: 14, italic: true, color: text, fontFace: theme.fontTitle, valign: "top" });
    });
    slide.addShape("roundRect", { x: 0.8, y: 5.55, w: 9.1, h: 1.1, fill: { color: mix(theme.primary, theme.background, 0.16) }, line: { type: "none" } as never, rectRadius: 0.1 });
    slide.addText(
      [{ text: "L'ASSISTANT    ", options: { bold: true, fontSize: 10, color: mix(theme.primary, theme.text, 0.62) } }, ...args.features.map((f) => ({ text: `\u2713 ${plain(f)}     `, options: { fontSize: 12 } }))],
      { x: 1.0, y: 5.6, w: 8.7, h: 1.0, valign: "middle", color: text, fontFace: theme.fontText },
    );
    const mascot = pptxPicture({ mascot: args.mascot }, { x: 10.2, y: 2.1, w: 2.4, h: 4.7 });
    if (mascot) slide.addImage(mascot);
  },
});
