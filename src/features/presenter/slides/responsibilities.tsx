import { z } from "zod";
import { pptxImage } from "./images";
import { Icon } from "./picture";
import { Frame, M, deep, defineTemplate, ease, faint, hex, mix, pptxTitle, rise, stagger, tint } from "./kit";

/** Who does what: a column per user under a header on the accent, its responsibilities as rows (an icon and a title each). */
export const responsibilities = defineTemplate({
  id: "responsibilities",
  description: "Who does what: 1-4 users (a name, an optional subtitle, a lucide icon), each with 1-6 responsibilities (a short title and a lucide icon), a column per user.",
  args: z.object({
  title: z.string(),
  users: z
    .array(
      z.object({
        name: z.string().describe("The user or role, e.g. Data owner"),
        subtitle: z.string().optional().describe("Who it is, in a few words"),
        icon: z.string().optional(),
        items: z.array(z.object({ title: z.string(), icon: z.string().optional() })).min(1).max(6),
      }),
    )
    .min(1)
    .max(4),
}),
  example: {
  title: "Qui fait quoi",
  users: [
    {
      name: "Data owner",
      subtitle: "Métier",
      icon: "crown",
      items: [
        { title: "Valide les définitions", icon: "badge-check" },
        { title: "Priorise les besoins", icon: "list-ordered" },
        { title: "Donne les accès", icon: "key-round" },
        { title: "Arbitre les écarts", icon: "scale" },
      ],
    },
    {
      name: "Key user",
      subtitle: "Relais des équipes",
      icon: "user-round-check",
      items: [
        { title: "Forme les utilisateurs", icon: "graduation-cap" },
        { title: "Publie les rapports", icon: "send" },
        { title: "Remonte les anomalies", icon: "bug" },
        { title: "Rédige les guides", icon: "book-open" },
      ],
    },
    {
      name: "Équipe data",
      subtitle: "BI & IT",
      icon: "database",
      items: [
        { title: "Maintient les modèles", icon: "boxes" },
        { title: "Surveille les refresh", icon: "refresh-cw" },
        { title: "Sécurise les données", icon: "shield-check" },
        { title: "Fait évoluer l'outil", icon: "rocket" },
      ],
    },
  ],
},
  Component: ({ args }) => (
    <Frame title={args.title}>
      <M.div variants={stagger(0.18, 0.2)} className="grid h-full content-center items-stretch gap-[2cqw]" style={{ gridTemplateColumns: `repeat(${args.users.length}, minmax(0, 1fr))` }}>
        {args.users.map((u, i) => (
          <M.div key={i} variants={rise} className="flex flex-col overflow-hidden rounded-[1.2cqw] border shadow-[0_0.3cqw_1cqw_rgba(0,0,0,0.06)]" style={{ borderColor: faint(10) }}>
            <div className="relative flex items-center gap-[1.2cqw] px-[1.8cqw] py-[1.6cqw]" style={{ background: "var(--s-primary)", color: "white" }}>
              <span className="grid size-[3.6cqw] shrink-0 place-items-center rounded-full" style={{ background: "rgba(255,255,255,0.2)" }}>
                <Icon name={u.icon ?? "user-round"} style={{ width: "1.9cqw", height: "1.9cqw" }} />
              </span>
              <span className="flex min-w-0 flex-col">
                <span className="text-[1.8cqw] leading-tight font-bold">{u.name}</span>
                {u.subtitle && <span className="text-[1.15cqw] opacity-80">{u.subtitle}</span>}
              </span>
            </div>
            <div className="flex flex-col py-[0.6cqw]">
              {u.items.map((it, j) => (
                <M.div key={j} initial={{ opacity: 0, x: "-1cqw" }} animate={{ opacity: 1, x: 0 }} transition={{ delay: 0.6 + i * 0.18 + j * 0.08, duration: 0.5, ease }} className="flex items-center gap-[1.2cqw] px-[1.8cqw] py-[1.05cqw] transition-colors hover:bg-[color-mix(in_srgb,var(--s-primary)_10%,transparent)]">
                  <span className="grid size-[3.2cqw] shrink-0 place-items-center rounded-[0.7cqw]" style={{ background: tint(20), color: deep }}>
                    <Icon name={it.icon} style={{ width: "1.7cqw", height: "1.7cqw" }} />
                  </span>
                  <span className="text-[1.5cqw] leading-snug">{it.title}</span>
                </M.div>
              ))}
            </div>
          </M.div>
        ))}
      </M.div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const text = hex(theme.text);
    const gap = 0.3;
    const w = (11.7 - gap * (args.users.length - 1)) / args.users.length;
    const rows = Math.max(...args.users.map((u) => u.items.length));
    const row = Math.min(0.55, 3.7 / rows);
    args.users.forEach((u, i) => {
      const x = 0.8 + i * (w + gap);
      const y = 2.1;
      slide.addShape("roundRect", { x, y, w, h: 1.05 + u.items.length * row + 0.15, fill: { color: hex(theme.background) }, line: { color: mix(theme.text, theme.background, 0.1), width: 0.75 }, rectRadius: 0.1 });
      slide.addShape("rect", { x, y, w, h: 0.95, fill: { color: hex(theme.primary) }, line: { type: "none" } as never });
      slide.addShape("ellipse", { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fill: { color: "FFFFFF", transparency: 78 }, line: { type: "none" } as never });
      const avatar = pptxImage(`icon-white:${u.icon?.trim() || "user-round"}`, { x: x + 0.33, y: y + 0.33, w: 0.29, h: 0.29 });
      if (avatar) slide.addImage(avatar);
      slide.addText(
        [{ text: u.name, options: { bold: true, fontSize: 15, breakLine: !!u.subtitle } }, ...(u.subtitle ? [{ text: u.subtitle, options: { fontSize: 10 } }] : [])],
        { x: x + 0.9, y: y + 0.1, w: w - 1.0, h: 0.75, color: "FFFFFF", fontFace: theme.fontText, valign: "middle" },
      );
      u.items.forEach((it, j) => {
        const ry = y + 1.1 + j * row;
        slide.addShape("roundRect", { x: x + 0.2, y: ry + (row - 0.38) / 2, w: 0.38, h: 0.38, fill: { color: mix(theme.primary, theme.background, 0.2) }, line: { type: "none" } as never, rectRadius: 0.2 });
        const icon = pptxImage(`icon:${it.icon?.trim() || "sparkles"}`, { x: x + 0.27, y: ry + (row - 0.24) / 2, w: 0.24, h: 0.24 });
        if (icon) slide.addImage(icon);
        slide.addText(it.title, { x: x + 0.7, y: ry, w: w - 0.85, h: row, fontSize: 12, color: text, fontFace: theme.fontText, valign: "middle" });
      });
    });
  },
});
