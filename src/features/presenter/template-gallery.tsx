/**
 * The Template tab: every layout of the company template as PowerPoint draws it, what it is for, and its
 * markdown — to copy, or to add to the presentation as a new slide (also to an imported PowerPoint).
 */
import { useEffect, useState } from "react";
import { CheckIcon, CopyIcon, Loader2Icon, PlusIcon } from "lucide-react";
import { toast } from "sonner";
import { LAYOUTS, type Layout } from "@shared/deck";
import { ErrorBox, Loading, Pills } from "@/components/common";
import { Button } from "@/components/ui/button";
import { errorText, fileUrl, get } from "@/lib/api";

type Example = { layout: Layout; title: string; group: string; markdown: string; picture?: string };
let loaded: Promise<{ examples: Example[] }> | undefined;

export function TemplateGallery({ onAdd }: { onAdd: (markdown: string) => Promise<void> }) {
  const [examples, setExamples] = useState<Example[]>();
  const [error, setError] = useState<string>();
  const [group, setGroup] = useState("Template");
  const [open, setOpen] = useState<number>();
  const [busy, setBusy] = useState<number>();
  const [copied, setCopied] = useState<number>();
  useEffect(() => {
    loaded ??= get("/deck/templates");
    loaded.then((r) => setExamples(r.examples), (e) => ((loaded = undefined), setError(errorText(e))));
  }, []);
  if (error) return <div className="p-6"><ErrorBox>{error}</ErrorBox></div>;
  if (!examples) return <Loading label="Drawing the template's layouts (the first time takes a little while)…" />;
  const groups = [...new Set(examples.map((e) => e.group))];
  return (
    <div className="flex flex-col gap-5 p-6">
      <div className="flex flex-wrap items-center gap-3">
        <Pills value={group} onChange={setGroup} options={groups.map((g) => ({ id: g, label: g, count: examples.filter((e) => e.group === g).length }))} />
        <div className="flex-1 text-right text-[12.5px] text-muted-foreground">Pictures in any layout: public links (`image:` or `![](https://…)`). Ask the assistant, or add a layout yourself.</div>
      </div>
      <div className="grid gap-5 lg:grid-cols-2 2xl:grid-cols-3">
        {examples.map((e, i) =>
          e.group !== group ? null : (
            <div key={i} className="flex flex-col overflow-hidden rounded-2xl border bg-card">
              {e.picture ? <img src={fileUrl(e.picture)} alt={e.title} className="w-full border-b" /> : <div className="aspect-video bg-muted" />}
              <div className="flex flex-1 flex-col gap-2 p-4">
                <div className="flex items-center gap-2">
                  <span className="flex-1 text-[14px] font-semibold">{e.title}</span>
                  <span className="rounded-full bg-muted px-2 py-0.5 font-mono text-[11px]">{e.layout}</span>
                </div>
                <div className="text-[12.5px] text-muted-foreground">{LAYOUTS[e.layout]}</div>
                {open === i && <pre className="max-h-56 overflow-auto rounded-xl bg-muted/60 p-3 font-mono text-[11.5px] whitespace-pre-wrap">{e.markdown}</pre>}
                <div className="mt-auto flex gap-2 pt-1">
                  <Button variant="ghost" size="sm" onClick={() => setOpen(open === i ? undefined : i)}>{open === i ? "Hide" : "Markdown"}</Button>
                  <Button variant="ghost" size="sm" onClick={() => void navigator.clipboard.writeText(e.markdown).then(() => (setCopied(i), setTimeout(() => setCopied(undefined), 1500)))}>{copied === i ? <CheckIcon /> : <CopyIcon />} Copy</Button>
                  <div className="flex-1" />
                  <Button variant="dark" size="sm" disabled={busy !== undefined} onClick={() => (setBusy(i), onAdd(e.markdown).then(() => toast.success("Slide added at the end"), (x) => toast.error(errorText(x))).finally(() => setBusy(undefined)))}>
                    {busy === i ? <Loader2Icon className="animate-spin" /> : <PlusIcon />} Add
                  </Button>
                </div>
              </div>
            </div>
          ),
        )}
      </div>
    </div>
  );
}
