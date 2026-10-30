import { CheckIcon } from "lucide-react";
import { Mascot } from "@/features/mascot";
import { pptxPicture } from "./picture";
import { Frame, M, Rich, deep, defineTemplate, ease, faint, hex, mix, plain, pptxTitle, tint } from "./kit";
import { Avatar, USECASE_EXAMPLE, pptxAvatar, usecaseArgs } from "./usecase";

/** Three dots that blink, then give way to the message. */
function Typing({ delay }: { delay: number }) {
  return (
    <M.span initial={{ opacity: 0 }} animate={{ opacity: [0, 1, 1, 0] }} transition={{ delay, duration: 0.7, times: [0, 0.1, 0.8, 1] }} className="absolute top-1/2 left-[6cqw] flex -translate-y-1/2 gap-[0.35cqw]">
      {[0, 1, 2].map((d) => (
        <M.span key={d} className="size-[0.6cqw] rounded-full" style={{ background: faint(40) }} animate={{ y: [0, "-0.3cqw", 0] }} transition={{ repeat: 2, duration: 0.3, delay: delay + d * 0.08 }} />
      ))}
    </M.span>
  );
}

/** A use case told as a conversation: each persona types its question (dots, then the bubble), the assistant answers with what it does. */
export const conversation = defineTemplate({
  id: "conversation",
  description: "A use case as a chat: each persona (1-3: role, a typical question, a lucide icon) asks its question in a bubble, the assistant answers in one bubble (an optional line + 1-5 features), the mascot (outfit) on the right. Same args as usecase.",
  args: usecaseArgs,
  example: USECASE_EXAMPLE,
  Component: ({ args }) => {
    const step = 0.9;
    const answer = 0.3 + args.personas.length * step;
    return (
      <Frame title={args.title}>
        <div className="grid h-full grid-cols-[66fr_34fr] gap-[2.4cqw]">
          <div className="flex min-h-0 flex-col justify-center gap-[1.2cqw]">
            {args.personas.map((p, i) => (
              <div key={i} className="relative flex items-end gap-[1cqw]">
                <M.span initial={{ opacity: 0, scale: 0.6 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.3 + i * step, duration: 0.4, ease }}>
                  <Avatar icon={p.icon} size={3.8} />
                </M.span>
                <Typing delay={0.3 + i * step} />
                <M.div initial={{ opacity: 0, y: "0.8cqw", scale: 0.96 }} animate={{ opacity: 1, y: 0, scale: 1 }} transition={{ delay: 0.3 + i * step + 0.6, duration: 0.5, ease }} className="flex max-w-[80%] flex-col gap-[0.3cqw] rounded-[1.4cqw] rounded-bl-[0.3cqw] px-[1.5cqw] py-[1cqw]" style={{ background: faint(6) }}>
                  <span className="text-[0.95cqw] font-semibold tracking-wide uppercase" style={{ color: deep }}>
                    {p.role}
                  </span>
                  <span className="text-[1.45cqw] leading-snug">
                    <Rich text={p.quote} />
                  </span>
                </M.div>
              </div>
            ))}
            <M.div initial={{ opacity: 0, y: "1cqw" }} animate={{ opacity: 1, y: 0 }} transition={{ delay: answer, duration: 0.6, ease }} className="ml-auto flex w-[78%] flex-col gap-[0.6cqw] rounded-[1.4cqw] rounded-br-[0.3cqw] px-[1.6cqw] py-[1.2cqw]" style={{ background: tint(28) }}>
              {args.description && (
                <span className="text-[1.25cqw] font-semibold leading-snug">
                  <Rich text={args.description} />
                </span>
              )}
              {args.features.map((f, i) => (
                <M.span key={i} initial={{ opacity: 0, x: "-0.8cqw" }} animate={{ opacity: 1, x: 0 }} transition={{ delay: answer + 0.3 + i * 0.12, duration: 0.4 }} className="flex items-center gap-[0.6cqw] text-[1.2cqw] leading-snug">
                  <span className="grid size-[1.5cqw] shrink-0 place-items-center rounded-full text-white" style={{ background: deep }}>
                    <CheckIcon style={{ width: "1cqw", height: "1cqw" }} strokeWidth={3} />
                  </span>
                  <Rich text={f} />
                </M.span>
              ))}
            </M.div>
          </div>
          <M.div initial={{ opacity: 0, x: "2cqw" }} animate={{ opacity: 1, x: 0 }} transition={{ delay: answer - 0.3, duration: 0.8, ease }} className="relative flex min-h-0 items-end justify-center">
            <span className="absolute bottom-[1cqw] h-[3cqw] w-[70%] rounded-[50%]" style={{ background: tint(30), filter: "blur(0.6cqw)" }} />
            <Mascot outfit={args.mascot} state="talking" size={400} className="relative h-auto max-h-full w-full" />
          </M.div>
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const text = hex(theme.text);
    const step = 1.05;
    args.personas.forEach((p, i) => {
      const y = 2.1 + i * step;
      pptxAvatar(slide, p.icon, { x: 0.8, y: y + 0.3, d: 0.5 }, ctx);
      slide.addShape("roundRect", { x: 1.45, y, w: 6.4, h: 0.9, fill: { color: mix(theme.text, theme.background, 0.06) }, line: { type: "none" } as never, rectRadius: 0.15 });
      slide.addText(
        [
          { text: p.role.toUpperCase(), options: { bold: true, fontSize: 9, color: mix(theme.primary, theme.text, 0.62), breakLine: true } },
          { text: plain(p.quote), options: { fontSize: 13 } },
        ],
        { x: 1.6, y, w: 6.1, h: 0.9, valign: "middle", color: text, fontFace: theme.fontText },
      );
    });
    const y = 2.2 + args.personas.length * step;
    const h = 0.3 + (args.description ? 0.4 : 0) + args.features.length * 0.32;
    slide.addShape("roundRect", { x: 3.0, y, w: 6.9, h, fill: { color: mix(theme.primary, theme.background, 0.28) }, line: { type: "none" } as never, rectRadius: 0.15 });
    slide.addText(
      [
        ...(args.description ? [{ text: plain(args.description), options: { bold: true, fontSize: 12, breakLine: true } }] : []),
        ...args.features.map((f, i) => ({ text: `\u2713  ${plain(f)}`, options: { fontSize: 12, breakLine: i < args.features.length - 1 } })),
      ],
      { x: 3.2, y, w: 6.5, h, valign: "middle", color: text, fontFace: theme.fontText },
    );
    const mascot = pptxPicture({ mascot: args.mascot }, { x: 10.2, y: 2.1, w: 2.4, h: 4.7 });
    if (mascot) slide.addImage(mascot);
  },
});
