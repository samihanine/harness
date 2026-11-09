import { Link, createFileRoute } from "@tanstack/react-router";
import { ScreenLogo } from "@/components/screen-logo";
import { APP_TITLE, useConfig } from "@/config";
import { shownFeatures, type Feature } from "@/lib/features";

/** "/": the home page — a card per screen shown (the root layout opens the first screen when it is off). */
export const Route = createFileRoute("/")({ component: Home });

const GROUP_TITLES = ["Assistants", "Library"];

function Home() {
  useConfig();
  return (
    <div className="h-full overflow-y-auto bg-muted/30">
      <div className="mx-auto flex max-w-6xl flex-col gap-8 p-8">
        <h1 className="text-[22px] font-semibold">{APP_TITLE}</h1>
        {([0, 1] as const).map((i) => {
          const shown = shownFeatures().filter((f) => f.group === i);
          if (!shown.length) return null;
          return (
            <section key={i} className="flex flex-col gap-3">
              <h2 className="text-[12.5px] font-medium tracking-wide text-muted-foreground uppercase">{GROUP_TITLES[i]}</h2>
              <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                {shown.map((f) => (
                  <Card key={f.id} feature={f} />
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
function Card({ feature }: { feature: Feature }) {
  return (
    <Link
      to={feature.path}
      className="flex flex-col items-center gap-3 rounded-2xl border bg-card p-6 text-center shadow-xs transition-all outline-none hover:-translate-y-0.5 hover:border-primary/40 hover:shadow-md focus-visible:ring-[3px] focus-visible:ring-ring/30"
    >
      <span className="flex h-36 items-center justify-center">
        <ScreenLogo feature={feature} size={"outfit" in feature.picture ? 128 : 80} live awake />
      </span>
      <span className="text-[16px] font-semibold">{feature.label}</span>
      <span className="text-[13px] leading-relaxed text-muted-foreground">{feature.description}</span>
    </Link>
  );
}
