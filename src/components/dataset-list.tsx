import { Link } from "@tanstack/react-router";
import { DatabaseIcon } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Switch } from "@/components/ui/switch";
import { Section } from "@/components/workspace";
import type { DatasetEntry } from "@/lib/store";

/**
 * The semantic models a screen's AI may query (DAX), one line each with its switch: `locked` ones
 * (e.g. the report's own model) are on and greyed. Without `onToggle`, a plain list.
 */
export function DatasetList({ title = "Semantic models", description, datasets, chosen, locked = [], onToggle }: { title?: string; description: string; datasets: DatasetEntry[]; chosen: string[]; locked?: string[]; onToggle?: (id: string, on: boolean) => void }) {
  return (
    <div className="mx-auto flex w-full max-w-4xl flex-col gap-4 overflow-y-auto p-6">
      <Section title={title} description={description}>
        {!datasets.length && <p className="text-[13px] text-muted-foreground">No semantic model in the app: add one in the Datasets screen.</p>}
        <div className="flex flex-col divide-y">
          {datasets.map((d) => {
            const on = locked.includes(d.id) || chosen.includes(d.id);
            return (
              <div key={d.id} className="flex items-center gap-3 py-2.5 text-[13.5px]">
                {onToggle && <Switch checked={on} disabled={locked.includes(d.id) || d.access === "view"} onCheckedChange={(v) => onToggle(d.id, v)} aria-label={`Use ${d.name}`} />}
                <Link to="/datasets" search={{ dataset: d.id }} className={`min-w-0 flex-1 truncate font-medium hover:underline ${locked.includes(d.id) ? "opacity-70" : ""}`}>
                  {d.name}
                </Link>
                {d.info && <span className="text-[12px] text-muted-foreground">{d.info.tables.filter((t) => !t.hidden).length} tables</span>}
                <Badge variant={d.access === "view" ? "outline" : "secondary"}>{d.access === "view" ? "no DAX access" : (d.access ?? "?")}</Badge>
                {locked.includes(d.id) && <span className="flex items-center gap-1 text-[12px] text-muted-foreground"><DatabaseIcon className="size-3.5" /> the report's</span>}
              </div>
            );
          })}
        </div>
      </Section>
    </div>
  );
}
