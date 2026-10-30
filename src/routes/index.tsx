import { Link, createFileRoute } from "@tanstack/react-router";
import { ScreenLogo } from "@/components/screen-logo";
import { APP_TITLE, SCREENS, SCREEN_DESCRIPTIONS, SCREEN_GROUPS, SCREEN_LABELS, SCREEN_PATHS, type Screen } from "@/config";

/** "/": the home page — a card per screen shown (the root layout opens the first screen when it is off). */
export const Route = createFileRoute("/")({ component: Home });

const GROUP_TITLES = ["Assistants", "Library"];

function Home() {
  return (
    <div className="h-full overflow-y-auto bg-muted/30">
      <div className="mx-auto flex max-w-6xl flex-col gap-8 p-8">
        <h1 className="text-[22px] font-semibold">{APP_TITLE}</h1>
        {SCREEN_GROUPS.map((group, i) => {
          const shown = group.filter((k) => SCREENS[k]);
          if (!shown.length) return null;
          return (
            <section key={i} className="flex flex-col gap-3">
              <h2 className="text-[12.5px] font-medium tracking-wide text-muted-foreground uppercase">{GROUP_TITLES[i]}</h2>
              <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                {shown.map((k) => (
                  <Card key={k} screen={k} />
                ))}
              </div>
            </section>
          );
        })}
      </div>
    </div>
  );
}

/** A screen: its picture (the mascot always awake here), name and purpose. */
function Card({ screen }: { screen: Screen }) {
  return (
    <Link
      to={SCREEN_PATHS[screen]}
      className="flex flex-col items-center gap-3 rounded-2xl border bg-card p-6 text-center shadow-xs transition-all outline-none hover:-translate-y-0.5 hover:border-primary/40 hover:shadow-md focus-visible:ring-[3px] focus-visible:ring-ring/30"
    >
      <span className="flex h-36 items-center justify-center">
        <ScreenLogo screen={screen} size={screen in { viewer: 1, editor: 1, research: 1, slides: 1 } ? 128 : 80} live awake />
      </span>
      <span className="text-[16px] font-semibold">{SCREEN_LABELS[screen]}</span>
      <span className="text-[13px] leading-relaxed text-muted-foreground">{SCREEN_DESCRIPTIONS[screen]}</span>
    </Link>
  );
}
