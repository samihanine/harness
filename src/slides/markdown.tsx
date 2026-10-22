import { z } from "zod";
import { Markdown } from "@/components/markdown";
import { Frame, M, defineTemplate, hex, pptxTitle, rise } from "./kit";

/** Free text (markdown: paragraphs, lists, bold), for what no other template fits. */
export const markdown = defineTemplate({
  id: "markdown",
  description: "Free markdown text (paragraphs, lists, **bold**) under a title, when no other template fits.",
  args: z.object({ title: z.string(), eyebrow: z.string().optional(), text: z.string() }),
  Component: ({ args }) => (
    <Frame eyebrow={args.eyebrow} title={args.title}>
      <M.div variants={rise} className="slide-markdown max-w-[70cqw] text-[1.5cqw] leading-relaxed">
        <Markdown>{args.text}</Markdown>
      </M.div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx, args.eyebrow);
    const lines = args.text.split("\n").filter((l) => l.trim());
    slide.addText(
      lines.map((l) => {
        const bullet = /^\s*[-*]\s+/.test(l);
        return { text: l.replace(/^\s*[-*]\s+/, "").replace(/^#+\s*/, "").replace(/\*\*([^*]+)\*\*/g, "$1"), options: { bullet, breakLine: true, bold: /^#/.test(l) } };
      }),
      { x: 0.8, y: 2.2, w: 11.5, h: 4.6, fontSize: 16, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), valign: "top", paraSpaceAfter: 6 },
    );
  },
});
