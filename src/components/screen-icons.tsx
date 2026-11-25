/**
 * The screens' icons in the rail: drawn in the colors of the tool each screen works with (Power BI's
 * yellow bars, Excel's green grid, PowerPoint's orange), animated when the mouse is over them (`group`
 * on the parent). Inspired by the products' logos, not copies of them.
 */
import type { ScreenId } from "@shared/types";
import { D_PATH, D_VIEWBOX } from "./dior-d";

const Bars = () => (
  <g>
    <rect x="3" y="10" width="4" height="10" rx="1.2" fill="#F2C811" className="origin-[5px_20px] transition-transform duration-300 group-hover:scale-y-75" />
    <rect x="8.5" y="6" width="4" height="14" rx="1.2" fill="#E8B10C" className="origin-[10.5px_20px] transition-transform delay-75 duration-300 group-hover:scale-y-110" />
    <rect x="14" y="2.5" width="4" height="17.5" rx="1.2" fill="#C99700" className="origin-[16px_20px] transition-transform delay-150 duration-300 group-hover:scale-y-90" />
  </g>
);

export function ScreenIcon({ screen, className = "size-[26px]" }: { screen: ScreenId; className?: string }) {
  switch (screen) {
    case "trainer":
      return (
        <svg viewBox="0 0 24 24" className={className} aria-hidden>
          <Bars />
          <g className="origin-[18px_18px] transition-transform duration-300 group-hover:scale-110">
            <circle cx="18" cy="18" r="5.6" fill="#fff" />
            <path d="M13.3 18c1.2-2.1 2.8-3.2 4.7-3.2s3.5 1.1 4.7 3.2c-1.2 2.1-2.8 3.2-4.7 3.2s-3.5-1.1-4.7-3.2Z" fill="#1F2937" className="origin-[18px_18px] group-hover:animate-[blink_0.6s_ease-in-out]" />
            <circle cx="18" cy="18" r="1.4" fill="#F2C811" />
          </g>
        </svg>
      );
    case "builder":
      return (
        <svg viewBox="0 0 24 24" className={className} aria-hidden>
          <Bars />
          <g className="origin-[18px_18px] transition-transform duration-500 group-hover:-rotate-[25deg]">
            <circle cx="18" cy="18" r="5.6" fill="#fff" />
            <path d="M20.6 14.6a2.6 2.6 0 0 0-3.4 3.3l-2.8 2.8a.9.9 0 0 0 1.3 1.3l2.8-2.8a2.6 2.6 0 0 0 3.3-3.4l-1.4 1.4-1.2-.2-.2-1.2 1.6-1.2Z" fill="#1F2937" />
          </g>
        </svg>
      );
    case "scientist":
      return (
        <svg viewBox="0 0 24 24" className={className} aria-hidden>
          <rect x="2.5" y="2.5" width="19" height="19" rx="4" fill="#107C41" />
          <g stroke="#fff" strokeOpacity=".35" strokeWidth="1" className="transition-opacity duration-300 group-hover:opacity-0">
            <path d="M12 5v14M5 9.5h14M5 14.5h14" />
          </g>
          <path d="M8 7.5l8 9M16 7.5l-8 9" stroke="#fff" strokeWidth="2.4" strokeLinecap="round" className="origin-center transition-transform duration-300 group-hover:rotate-90 group-hover:scale-90" />
        </svg>
      );
    case "presenter":
      return (
        <svg viewBox="0 0 24 24" className={className} aria-hidden>
          <rect x="2.5" y="2.5" width="19" height="19" rx="4" fill="#C43E1C" />
          <circle cx="16.5" cy="8" r="3.2" fill="#F4A582" className="origin-[16.5px_8px] transition-transform duration-300 group-hover:scale-125" />
          <path d="M8 18V7h3.6a3 3 0 0 1 0 6H8" fill="none" stroke="#fff" strokeWidth="2.3" strokeLinecap="round" strokeLinejoin="round" className="transition-transform duration-300 group-hover:-translate-y-0.5" />
        </svg>
      );
  }
}

/** The app's logo: the D of the Dior wordmark. */
export const DiorD = ({ className }: { className?: string }) => (
  <svg viewBox={D_VIEWBOX} className={className} aria-hidden>
    <path d={D_PATH} fill="currentColor" fillRule="evenodd" />
  </svg>
);
