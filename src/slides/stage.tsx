/**
 * Drawing slides: one slide in a 16:9 box (container units), animated in when shown, and the theme as
 * CSS variables (--s-primary, --s-bg, --s-text, --s-title, --s-text-font). Fonts are Google fonts.
 */
import { useEffect, type CSSProperties } from "react";
import { MotionConfig, motion } from "motion/react";
import { checkSlide, type Deck, type SlideJson } from "./index";
import type { DeckInfo, Theme } from "./kit";

const SYSTEM = ["Plus Jakarta Sans", "system-ui", "Arial", "Helvetica", "Georgia"];

/** Loads the deck's fonts from Google Fonts (once each). */
export function useFonts(theme: Theme) {
  useEffect(() => {
    for (const font of [theme.fontTitle, theme.fontText]) {
      if (!font || SYSTEM.includes(font)) continue;
      const id = `font-${font.replace(/\W+/g, "-")}`;
      if (document.getElementById(id)) continue;
      document.head.append(Object.assign(document.createElement("link"), { id, rel: "stylesheet", href: `https://fonts.googleapis.com/css2?family=${encodeURIComponent(font)}:wght@400;500;600;700&display=swap` }));
    }
  }, [theme.fontTitle, theme.fontText]);
}

export const themeStyle = (theme: Theme): CSSProperties =>
  ({ "--s-primary": theme.primary, "--s-bg": theme.background, "--s-text": theme.text, "--s-title": `"${theme.fontTitle}", system-ui, sans-serif`, "--s-text-font": `"${theme.fontText}", system-ui, sans-serif`, background: theme.background, color: theme.text, fontFamily: "var(--s-text-font)" }) as CSSProperties;

/** One slide, drawn at the box's size; `still`: shown at once (thumbnails, print). */
export function SlideView({ slide, deck, info, still = false }: { slide: SlideJson; deck: Deck; info: DeckInfo; still?: boolean }) {
  let content;
  let bare = false;
  try {
    const { template, args } = checkSlide(slide);
    bare = !!template.bare;
    content = <template.Component args={args} theme={deck.theme} deck={info} />;
  } catch (e) {
    content = <div className="flex size-full items-center justify-center p-[4cqw] text-center text-[1.4cqw] text-red-600">{e instanceof Error ? e.message : String(e)}</div>;
  }
  return (
    <div className="relative aspect-video w-full overflow-hidden [container-type:inline-size]" style={themeStyle(deck.theme)}>
      <MotionConfig reducedMotion={still ? "always" : "user"}>
        <motion.div className="size-full" initial={still ? "show" : "hidden"} animate="show">
          {content}
        </motion.div>
      </MotionConfig>
      {!bare && (
        <div className="pointer-events-none absolute inset-x-[6cqw] bottom-[2cqw] flex justify-between text-[0.95cqw] opacity-45">
          <span>{deck.title}{info.sections[info.section] ? ` · ${info.sections[info.section]}` : ""}</span>
          <span className="tabular-nums">{info.index + 1} / {info.count}</span>
        </div>
      )}
    </div>
  );
}
