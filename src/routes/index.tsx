import { Link, createFileRoute } from "@tanstack/react-router";
import { ArrowRightIcon } from "lucide-react";
import { SCREENS } from "@shared/types";
import { SCREEN_INFO } from "@/components/screens";
import { Mascot, outfitOf } from "@/features/mascot";
import { useSettings } from "@/lib/settings";

export const Route = createFileRoute("/")({ component: Home });

function Home() {
  const settings = useSettings();
  const shown = SCREENS.filter((s) => settings?.general.screens[s] !== false);
  const main = shown.filter((s) => ["trainer", "builder", "scientist", "presenter"].includes(s));
  const other = shown.filter((s) => !main.includes(s));
  return (
    <div className="h-full overflow-auto rounded-3xl border bg-card">
      <div className="mx-auto flex max-w-6xl flex-col gap-8 px-8 py-10">
        <div>
          <h1 className="text-[34px] font-light tracking-tight">Hub</h1>
          <p className="mt-1 text-[15px] text-muted-foreground">Power BI, Excel analyses and PowerPoint, with assistants that explore, ask, then act.</p>
        </div>
        <div className="grid gap-4 sm:grid-cols-2">
          {main.map((s) => {
            const info = SCREEN_INFO[s];
            return (
              <Link key={s} to={info.path} className="group flex items-center gap-5 rounded-3xl border bg-background/60 p-5 transition-all hover:-translate-y-0.5 hover:bg-card hover:shadow-[0_6px_24px_rgba(0,0,0,0.06)]">
                {settings?.general.mascot !== false ? <Mascot outfit={outfitOf(s)} state="idle" size={92} /> : <info.icon className="size-10" />}
                <div className="min-w-0 flex-1">
                  <div className="text-[12px] font-medium tracking-wide text-primary uppercase">{info.role}</div>
                  <div className="text-[20px] font-semibold">{info.label}</div>
                  <div className="mt-1 text-[13.5px] text-muted-foreground">{info.description}</div>
                </div>
                <ArrowRightIcon className="size-5 text-muted-foreground transition-transform group-hover:translate-x-1" />
              </Link>
            );
          })}
        </div>
        <div className="grid gap-3 sm:grid-cols-3">
          {other.map((s) => {
            const info = SCREEN_INFO[s];
            return (
              <Link key={s} to={info.path} className="flex items-start gap-3 rounded-2xl border p-4 hover:bg-muted/50">
                <info.icon className="mt-0.5 size-5 text-muted-foreground" />
                <div>
                  <div className="font-semibold">{info.label}</div>
                  <div className="text-[12.5px] text-muted-foreground">{info.description}</div>
                </div>
              </Link>
            );
          })}
        </div>
      </div>
    </div>
  );
}
