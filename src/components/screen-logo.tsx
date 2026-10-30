/**
 * The picture of each screen (sidebar, home page): the mascot dressed for the four assistants, the
 * logo of what the four library screens hold (Power BI reports, semantic models, Excel, documents).
 */
import { WakingMascot, Mascot, type Outfit } from "@/features/mascot";
import type { Screen } from "@/config";
import contextIcon from "@/assets/context-icon.png";

const OUTFIT: Partial<Record<Screen, Outfit>> = { viewer: "trainer", editor: "builder", research: "datascientist", slides: "presenter" };

/** Power BI: three yellow bars. */
export const PowerBiLogo = ({ size = 24 }: { size?: number }) => (
  <svg viewBox="0 0 32 32" width={size} height={size} aria-hidden>
    <rect x={5} y={16} width={7} height={12} rx={2} fill="#F2C811" />
    <rect x={12.5} y={10} width={7} height={18} rx={2} fill="#E8B10A" />
    <rect x={20} y={4} width={7} height={24} rx={2} fill="#C99700" />
  </svg>
);

/** A semantic model: stacked tables in the Power BI yellow. */
export const ModelLogo = ({ size = 24 }: { size?: number }) => (
  <svg viewBox="0 0 32 32" width={size} height={size} aria-hidden>
    <ellipse cx={16} cy={8} rx={11} ry={4} fill="#F2C811" />
    <path d="M5 8 L5 24 Q5 28 16 28 Q27 28 27 24 L27 8 Q27 12 16 12 Q5 12 5 8 Z" fill="#E8B10A" />
    <path d="M5 14 Q5 18 16 18 Q27 18 27 14 M5 20 Q5 24 16 24 Q27 24 27 20" fill="none" stroke="#C99700" strokeWidth={1.6} />
    <ellipse cx={16} cy={8} rx={11} ry={4} fill="none" stroke="#C99700" strokeWidth={1.2} />
  </svg>
);

/** Excel: the green sheet with its X. */
export const ExcelLogo = ({ size = 24 }: { size?: number }) => (
  <svg viewBox="0 0 32 32" width={size} height={size} aria-hidden>
    <rect x={10} y={4} width={19} height={24} rx={2} fill="#21A366" />
    <path d="M10 10 H29 M10 16 H29 M10 22 H29 M19.5 4 V28" stroke="#107C41" strokeWidth={1.2} />
    <rect x={3} y={8} width={15} height={16} rx={2} fill="#107C41" />
    <path d="M7 12 L14 20 M14 12 L7 20" stroke="#fff" strokeWidth={2.4} strokeLinecap="round" />
  </svg>
);

/** Contexts: a text document (icon by Flaticon, credited in the README). */
export const ContextLogo = ({ size = 24 }: { size?: number }) => <img src={contextIcon} width={size} height={size} alt="" aria-hidden />;

/** Confluence: a blue wiki page (generic, not the brand's logo). */
export const WikiLogo = ({ size = 24 }: { size?: number }) => (
  <svg viewBox="0 0 32 32" width={size} height={size} aria-hidden>
    <path d="M8 3 H20 L26 9 V27 Q26 29 24 29 H8 Q6 29 6 27 V5 Q6 3 8 3 Z" fill="#1868DB" />
    <path d="M20 3 V8 Q20 9 21 9 H26 Z" fill="#7AB3FF" />
    <path d="M10 14 H22 M10 18 H22 M10 22 H18" stroke="#fff" strokeWidth={2} strokeLinecap="round" />
  </svg>
);

const LOGOS: Partial<Record<Screen, typeof PowerBiLogo>> = { reports: PowerBiLogo, datasets: ModelLogo, excel: ExcelLogo, contexts: ContextLogo, confluence: WikiLogo };

/**
 * A screen's picture at `size` px. `live`: the mascot asleep, awake while hovered (or `awake`); else
 * standing still (the sidebar).
 */
export function ScreenLogo({ screen, size, live, awake }: { screen: Screen; size: number; live?: boolean; awake?: boolean }) {
  const outfit = OUTFIT[screen];
  if (outfit) return live ? <WakingMascot outfit={outfit} size={size} awake={awake} /> : <Mascot outfit={outfit} size={size} still />;
  const Logo = LOGOS[screen] ?? PowerBiLogo;
  return <Logo size={size} />;
}
