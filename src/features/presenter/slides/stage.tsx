/**
 * Drawing slides: one slide in a 16:9 box (container units), animated in when shown, and the theme as
 * CSS variables (--s-primary, --s-bg, --s-text, --s-title, --s-text-font). Fonts are Google fonts.
 */
import { useEffect, type CSSProperties } from "react";
import { MotionConfig, motion } from "motion/react";
import { checkSlide, type Deck, type SlideJson } from ".";
import { SlideContext, brandOf, type DeckInfo, type Theme } from "./kit";
import { MONOGRAM_FONT, Monogram } from "./monogram";
import { StillSlide } from "./powerbi";

const SYSTEM = ["Plus Jakarta Sans", "system-ui", "Arial", "Helvetica", "Georgia"];

/** Loads the deck's fonts from Google Fonts (once each). */
export function useFonts(theme: Theme) {
  useEffect(() => {
    // The monogram's Didone comes with the header that shows it.
    for (const font of [theme.fontTitle, theme.fontText, MONOGRAM_FONT]) {
      if (!font || SYSTEM.includes(font)) continue;
      const id = `font-${font.replace(/\W+/g, "-")}`;
      if (document.getElementById(id)) continue;
      document.head.append(Object.assign(document.createElement("link"), { id, rel: "stylesheet", href: `https://fonts.googleapis.com/css2?family=${encodeURIComponent(font)}:wght@400;500;600;700&display=swap` }));
    }
  }, [theme.fontTitle, theme.fontText, theme.header]);
}

export const themeStyle = (theme: Theme): CSSProperties =>
  ({ "--s-secondary": theme.secondary ?? "#8b5a2b", "--s-success": theme.success ?? "#15803d", "--s-brand": `"${MONOGRAM_FONT}", Didot, "Bodoni 72", serif`, "--s-primary": theme.primary, "--s-bg": theme.background, "--s-text": theme.text, "--s-title": `"${theme.fontTitle}", system-ui, sans-serif`, "--s-text-font": `"${theme.fontText}", system-ui, sans-serif`, background: theme.background, color: theme.text, fontFamily: "var(--s-text-font)" }) as CSSProperties;

/** One slide, drawn at the box's size; `still`: shown at once (thumbnails, print). */
export function SlideView({ slide, deck, info, still = false }: { slide: SlideJson; deck: Deck; info: DeckInfo; still?: boolean }) {
  let content;
  let bare = false;
  let page = false;
  let title = "";
  try {
    const { template, args } = checkSlide(slide);
    bare = !!template.bare;
    page = (template.kind ?? "content") === "content" && !bare;
    title = String(args.title ?? "");
    // The "brand" header shows the page title: the template leaves it out.
    content = (
      <SlideContext.Provider value={{ titleInHeader: page && deck.theme.header !== "none" }}>
        <template.Component args={args} theme={deck.theme} deck={info} />
      </SlideContext.Provider>
    );
  } catch (e) {
    content = <div className="flex size-full items-center justify-center p-[4cqw] text-center text-[1.4cqw] text-red-600">{e instanceof Error ? e.message : String(e)}</div>;
  }
  return (
    <div className={`relative aspect-video w-full overflow-hidden [container-type:inline-size] ${still ? "slide-still" : ""}`} style={themeStyle(deck.theme)}>
      {page && <Header deck={deck} info={info} still={still} title={title} />}
      <MotionConfig reducedMotion={still ? "always" : "user"}>
        <StillSlide.Provider value={still}>
          <motion.div className="size-full" initial={still ? "show" : "hidden"} animate="show">
            {content}
          </motion.div>
        </StillSlide.Provider>
      </MotionConfig>
      {!bare && (
        // Where the page is is said by the header: the footer keeps the page number.
        <div className="pointer-events-none absolute right-[6cqw] bottom-[2cqw] text-[0.95cqw] tabular-nums opacity-45">
          {info.index + 1} / {info.count}
        </div>
      )}
    </div>
  );
}

/**
 * The header of the content pages (deck's general config): the brand letter drawn as on the cover, the
 * page title on the left, and on the right where the page is (deck › part › subpart) or ("brand") the
 * brand name; "rule" and "brand" over a fine rule.
 */
function Header({ deck, info, still, title }: { deck: Deck; info: DeckInfo; still: boolean; title: string }) {
  const variant = deck.theme.header ?? "breadcrumb";
  if (variant === "none") return null;
  const trail = [deck.title, info.sections[info.section], info.subsection >= 0 ? info.subsections[info.subsection] : undefined].filter(Boolean) as string[];
  return (
    <div className="pointer-events-none absolute inset-x-[5cqw] top-[1.6cqw] z-10 flex flex-col gap-[0.7cqw]">
      <div className="flex h-[4cqw] items-center gap-[1.2cqw]">
        <Monogram letter={brandOf(deck.theme).charAt(0)} still={still} className="h-[3.4cqw] w-auto shrink-0" />
        <motion.span initial={still ? false : { opacity: 0, x: 12 }} animate={{ opacity: 1, x: 0 }} transition={{ duration: 0.8, delay: 0.15 }} className="min-w-0 flex-1 truncate text-[2.2cqw] leading-tight font-bold" style={{ fontFamily: "var(--s-title)" }}>
          {title}
          {variant === "brand" && info.group.count > 1 && <span className="font-medium opacity-50">{` (${info.group.index + 1}/${info.group.count})`}</span>}
        </motion.span>
        {variant === "brand" ? (
          <span className="shrink-0 text-[1.4cqw] italic opacity-60" style={{ fontFamily: "var(--s-brand)" }}>
            {brandOf(deck.theme)}
          </span>
        ) : (
          <span className="flex max-w-[45%] shrink-0 items-center gap-[0.6cqw] truncate text-[0.9cqw] leading-tight font-medium tracking-[0.14em] uppercase">
            {trail.map((t, i) => (
              <span key={i} className="flex items-center gap-[0.6cqw] truncate" style={{ opacity: i === trail.length - 1 ? 0.85 : 0.45, color: i === trail.length - 1 ? "var(--s-primary)" : undefined }}>
                {i > 0 && <span style={{ opacity: 0.5 }}>›</span>}
                {t}
              </span>
            ))}
          </span>
        )}
      </div>
      {(variant === "rule" || variant === "brand") && (
        <motion.div initial={still ? false : { scaleX: 0 }} animate={{ scaleX: 1 }} transition={{ duration: 1.1, ease: [0.22, 1, 0.36, 1], delay: 0.2 }} className="h-px w-full origin-left" style={{ background: "color-mix(in srgb, var(--s-text) 20%, transparent)" }} />
      )}
    </div>
  );
}
