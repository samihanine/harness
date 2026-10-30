/**
 * The mascot's outfits: each fills some slots of the body (back, feet, torso, sleeves, head top, eyes,
 * face, the prop in its right hand). The slots sit in the parts that move (the hat in the head, the
 * prop in the hand), so every animation works with every outfit.
 */
import type { ReactNode } from "react";

/** Colors of the drawing, taken from the reference (a cartoon hedgehog with turquoise diving gear). */
export const C = {
  line: "#17191b",
  fur: "#5f6468",
  furShade: "#4a4f53",
  suit: "#53585c",
  suitShade: "#3e4246",
  spikes: "#3b4043",
  spikeLight: "#767c80",
  white: "#ffffff",
  pink: "#f4a3a3",
  mouth: "#5a1f22",
  tongue: "#f47e86",
  /** The turquoise of the reference (diving gear, accents of every outfit). */
  accent: "#22c3c6",
  accentLight: "#a5efec",
  accentDark: "#14979a",
};

const S = { stroke: C.line, strokeWidth: 3.6, strokeLinejoin: "round" as const, strokeLinecap: "round" as const };

/** `hood`: behind the head and moving with it (over the quills). */
export type Slots = { hood?: ReactNode; back?: ReactNode; feet?: ReactNode; torso?: ReactNode; sleeve?: string; headTop?: ReactNode; eyes?: ReactNode; face?: ReactNode; prop?: ReactNode };

/** The body's outline (the drawing's), for clothes covering it. */
export const BODY = "M71 124 Q56 150 65 174 Q80 189 100 189 Q120 189 135 174 Q144 150 129 124 Q116 115 100 115 Q84 115 71 124 Z";

/** Plain shoes (outfits without their own feet). */
const shoes = (color: string) => (
  <>
    <path d="M74 199 Q74 189 87 189 Q98 189 98 196 Q98 202 88 202 L79 202 Q74 202 74 199 Z" fill={color} {...S} />
    <path d="M126 199 Q126 189 113 189 Q102 189 102 196 Q102 202 112 202 L121 202 Q126 202 126 199 Z" fill={color} {...S} />
  </>
);

/** A jacket over the body (its color) with the shirt in its V and what is at the neck. */
const jacket = (color: string, neck: ReactNode) => (
  <>
    <path d={BODY} fill={color} {...S} />
    <path d="M87 120 L100 152 L113 120 Q100 116 87 120 Z" fill={C.white} {...S} strokeWidth={2.4} />
    <path d="M100 152 L100 188" {...S} strokeWidth={2.2} />
    {neck}
  </>
);

/** Glasses over the eyes (round or square), with their bridge. */
const glasses = (square: boolean, tint = C.white) => (
  <>
    {square ? (
      <>
        <rect x={67} y={73} width={29} height={24} rx={4} fill={tint} fillOpacity={0.2} {...S} strokeWidth={3} />
        <rect x={105} y={73} width={29} height={24} rx={4} fill={tint} fillOpacity={0.2} {...S} strokeWidth={3} />
      </>
    ) : (
      <>
        <circle cx={82} cy={85} r={14} fill={tint} fillOpacity={0.15} {...S} strokeWidth={2.6} />
        <circle cx={119} cy={85} r={14} fill={tint} fillOpacity={0.15} {...S} strokeWidth={2.6} />
      </>
    )}
    <path d={square ? "M96 82 Q100.5 79 105 82" : "M96 83 Q100.5 80 105 83"} fill="none" {...S} strokeWidth={2.6} />
    <path d={square ? "M67 80 L56 77 M134 80 L145 77" : "M68 82 L56 78 M133 82 L145 78"} {...S} strokeWidth={2.6} />
  </>
);

export const OUTFITS = {
  diver: {
    // The tank on its back, seen on the left of the body.
    back: (
      <>
        <rect x={44} y={116} width={26} height={56} rx={11} fill={C.accent} {...S} />
        <rect x={49} y={122} width={6} height={42} rx={3} fill={C.accentLight} opacity={0.85} />
        <rect x={50} y={107} width={13} height={11} rx={3} fill={C.fur} {...S} />
      </>
    ),
    // Fins, spread as in the drawing.
    feet: (
      <>
        <path d="M82 188 L52 211 Q62 221 77 217 Q89 214 97 198 L97 190 Z" fill={C.accent} {...S} />
        <path d="M78 199 L63 213 M87 200 L76 216" stroke={C.accentDark} strokeWidth={2.6} strokeLinecap="round" />
        <path d="M118 188 L148 211 Q138 221 123 217 Q111 214 103 198 L103 190 Z" fill={C.accent} {...S} />
        <path d="M122 199 L137 213 M113 200 L124 216" stroke={C.accentDark} strokeWidth={2.6} strokeLinecap="round" />
      </>
    ),
    // The harness: two straps and the belt with its buckle.
    torso: (
      <>
        <path d="M76 124 L80 162 M124 124 L120 162" stroke={C.line} strokeWidth={12} strokeLinecap="round" />
        <path d="M76 124 L80 162 M124 124 L120 162" stroke={C.spikes} strokeWidth={6.5} strokeLinecap="round" />
        <rect x={63} y={158} width={74} height={13} rx={4} fill={C.spikes} {...S} />
        <rect x={69} y={161.5} width={8} height={6} rx={1} fill={C.fur} />
        <rect x={123} y={161.5} width={8} height={6} rx={1} fill={C.fur} />
        <rect x={90} y={155} width={20} height={19} rx={3} fill={C.fur} {...S} strokeWidth={2.8} />
        <rect x={95} y={159.5} width={10} height={10} rx={1.5} fill={C.spikes} />
      </>
    ),
    // The mask pushed up on the forehead, its strap around the head.
    headTop: (
      <>
        <path d="M50 58 Q100 44 150 58 L150 66 Q100 52 50 66 Z" fill={C.spikes} {...S} strokeWidth={2.6} />
        <path d="M64 50 Q64 34 82 34 L118 34 Q136 34 136 50 L136 58 Q136 70 122 70 L108 70 L100 62 L92 70 L78 70 Q64 70 64 58 Z" fill={C.accent} {...S} />
        <rect x={70} y={40} width={26} height={22} rx={8} fill={C.accentLight} stroke={C.accentDark} strokeWidth={1.6} />
        <rect x={104} y={40} width={26} height={22} rx={8} fill={C.accentLight} stroke={C.accentDark} strokeWidth={1.6} />
        <ellipse cx={78} cy={47} rx={3.2} ry={2.4} fill={C.white} />
        <ellipse cx={112} cy={47} rx={3.2} ry={2.4} fill={C.white} />
        <circle cx={87} cy={50} r={1.6} fill={C.white} />
        <circle cx={121} cy={50} r={1.6} fill={C.white} />
      </>
    ),
    // The snorkel: from the mouth, up along the right of the head.
    face: (
      <>
        <path d="M112 106 Q140 110 146 92 L150 24" fill="none" stroke={C.line} strokeWidth={13} strokeLinecap="round" />
        <path d="M112 106 Q140 110 146 92 L150 24" fill="none" stroke={C.accent} strokeWidth={7.5} strokeLinecap="round" />
        <path d="M148 26 Q146 12 158 10" fill="none" stroke={C.line} strokeWidth={13} strokeLinecap="round" />
        <path d="M148 26 Q146 12 158 10" fill="none" stroke={C.accent} strokeWidth={7.5} strokeLinecap="round" />
        <rect x={140} y={60} width={15} height={8} rx={2.5} fill={C.fur} {...S} strokeWidth={2.4} />
        <rect x={107} y={99} width={13} height={14} rx={4} fill={C.accent} {...S} strokeWidth={2.8} />
        <path d="M110 103 L117 103 M110 107 L117 107" stroke={C.accentDark} strokeWidth={1.6} />
      </>
    ),
  },
  /** The trainer, as a coach: a blue polo with its collar, a cap with its visor, a whistle, a clipboard. */
  trainer: {
    feet: shoes(C.white),
    sleeve: "#2563eb",
    torso: (
      <>
        <path d={BODY} fill="#2563eb" {...S} />
        {/* The polo's collar and its buttons. */}
        <path d="M84 119 L100 132 L116 119 L112 128 L100 138 L88 128 Z" fill={C.white} {...S} strokeWidth={2.4} />
        <path d="M100 138 L100 154" {...S} strokeWidth={2} />
        <circle cx={100} cy={145} r={1.6} fill={C.line} />
        {/* The whistle on its turquoise cord. */}
        <path d="M86 126 Q90 150 100 158 Q110 150 114 126" fill="none" stroke={C.accent} strokeWidth={2.6} />
        <path d="M93 156 L107 156 Q112 156 112 162 Q112 168 105 168 L98 168 Q92 168 92 163 L86 163 L86 158 Z" fill="#cbd5e1" {...S} strokeWidth={2.4} />
        <rect x={63} y={170} width={74} height={6} fill="#1f2937" />
      </>
    ),
    // A cap, its visor to the front, the brand's turquoise on top.
    headTop: (
      <>
        <path d="M60 60 Q60 26 100 26 Q140 26 140 60 Z" fill="#2563eb" {...S} />
        <path d="M100 26 L100 60 M80 30 Q74 44 74 60 M120 30 Q126 44 126 60" fill="none" stroke="#1e40af" strokeWidth={1.8} />
        <circle cx={100} cy={26} r={4} fill={C.accent} {...S} strokeWidth={2} />
        <path d="M56 58 Q100 50 144 58 Q148 66 136 68 Q100 62 64 68 Q52 66 56 58 Z" fill="#1d4ed8" {...S} />
      </>
    ),
    // A clipboard with its sheet and checks.
    prop: (
      <>
        <rect x={128} y={146} width={26} height={33} rx={3} fill="#a16207" {...S} />
        <rect x={132} y={152} width={18} height={23} rx={1} fill={C.white} />
        <rect x={136} y={143} width={10} height={6} rx={2} fill="#9ca3af" {...S} strokeWidth={2} />
        <path d="M134 158 L136 160 L139 156 M134 166 L136 168 L139 164" fill="none" stroke={C.accent} strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round" />
        <path d="M141 158 L148 158 M141 166 L148 166" stroke={C.line} strokeWidth={1.5} strokeLinecap="round" />
      </>
    ),
  },
  builder: {
    feet: shoes("#5a3a1e"),
    torso: (
      <>
        <path d={BODY} fill="#f97316" {...S} />
        <rect x={93} y={116} width={14} height={73} fill={C.fur} />
        <path d="M93 117 L93 189 M107 117 L107 189" {...S} strokeWidth={2.2} />
        <rect x={66} y={142} width={24} height={6} rx={2} fill="#e5e7eb" />
        <rect x={110} y={142} width={24} height={6} rx={2} fill="#e5e7eb" />
        <rect x={63} y={166} width={74} height={10} rx={3} fill="#8b5a2b" {...S} />
        <rect x={67} y={169} width={14} height={13} rx={2} fill="#a16207" {...S} strokeWidth={2.2} />
      </>
    ),
    headTop: (
      <>
        <path d="M56 62 Q56 20 100 20 Q144 20 144 62 Z" fill="#f5b400" {...S} />
        <rect x={94} y={23} width={12} height={37} rx={3} fill="#ffd23f" />
        <rect x={48} y={57} width={104} height={10} rx={5} fill="#e0a100" {...S} />
      </>
    ),
    prop: (
      <>
        <rect x={122} y={160} width={8} height={32} rx={3} fill="#9ca3af" {...S} />
        <rect x={113} y={188} width={26} height={11} rx={3} fill="#6b7280" {...S} />
      </>
    ),
  },
  /** The data scientist, as a developer: a purple hoodie with the hood up, headphones, square glasses, a laptop. */
  datascientist: {
    feet: shoes("#e5e7eb"),
    sleeve: "#5b3fa8",
    // The hood up behind the head.
    hood: <path d="M48 84 Q46 30 100 26 Q154 30 152 84 Q154 118 132 128 L68 128 Q46 118 48 84 Z" fill="#5b3fa8" {...S} />,
    torso: (
      <>
        <path d={BODY} fill="#5b3fa8" {...S} />
        {/* The hood's opening, its strings and the front pocket with the code sign. */}
        <path d="M80 122 Q100 138 120 122" fill="none" {...S} strokeWidth={2.6} />
        <path d="M93 130 L92 148 M107 130 L108 148" stroke={C.white} strokeWidth={2.4} strokeLinecap="round" />
        <path d="M78 166 Q78 154 88 154 L112 154 Q122 154 122 166 L122 176 L78 176 Z" fill="#4a3290" {...S} strokeWidth={2.4} />
        <text x={100} y={170} textAnchor="middle" fontSize={10} fontWeight={800} fill={C.accent} fontFamily="ui-monospace, monospace">{"</>"}</text>
      </>
    ),
    eyes: glasses(true, "#0f172a"),
    // Headphones over the head, their cushions on the ears.
    headTop: (
      <>
        <path d="M54 76 Q50 30 100 30 Q150 30 146 76" fill="none" stroke={C.line} strokeWidth={11} strokeLinecap="round" />
        <path d="M54 76 Q50 30 100 30 Q150 30 146 76" fill="none" stroke="#1f2937" strokeWidth={6} strokeLinecap="round" />
        <rect x={42} y={62} width={16} height={26} rx={6} fill={C.accent} {...S} />
        <rect x={142} y={62} width={16} height={26} rx={6} fill={C.accent} {...S} />
      </>
    ),
    // A laptop held open, its screen showing a chart.
    prop: (
      <>
        <path d="M110 150 L146 150 L146 174 L110 174 Z" fill="#cbd5e1" {...S} />
        <rect x={114} y={154} width={28} height={16} rx={1.5} fill="#0f172a" />
        <path d="M117 166 L123 161 L128 164 L136 157 L139 159" fill="none" stroke={C.accent} strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" />
        <path d="M106 176 L150 176 L146 181 L110 181 Z" fill="#94a3b8" {...S} strokeWidth={2.6} />
      </>
    ),
  },
  presenter: {
    feet: shoes("#1f2937"),
    sleeve: C.accent,
    torso: jacket(C.accent, <path d="M97 122 L103 122 L105 143 L100 150 L95 143 Z" fill={C.line} />),
    face: (
      <>
        <path d="M146 76 Q148 104 116 106" fill="none" stroke={C.line} strokeWidth={2.6} strokeLinecap="round" />
        <circle cx={115} cy={106} r={3.2} fill={C.line} />
      </>
    ),
    prop: (
      <>
        <rect x={122} y={154} width={8} height={24} rx={3} fill={C.line} />
        <circle cx={126} cy={152} r={4} fill={C.accent} {...S} strokeWidth={2} />
      </>
    ),
  },
} satisfies Record<string, Slots>;

export type Outfit = keyof typeof OUTFITS;

/** Outfits of older versions, by their new name (a deck may still name them). */
export const OUTFIT_ALIASES: Record<string, Outfit> = { teacher: "trainer", scientist: "datascientist" };
