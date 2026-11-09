import { useRef, useState } from "react";
import { DownloadIcon, UploadIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ScrollPage, Section } from "@/components/workspace";
import { useCollection } from "@/lib/store";
import { downloadGuides, importGuides } from "./guides";

/** Settings › Guides (given by the Trainer). */
export default function GuidesSettings() {
  return (
    <ScrollPage>
      <Guides />
    </ScrollPage>
  );
}

/** The guides of the report viewer, kept in this browser: download them as Excel, import Excel or CSV. */
function Guides() {
  const guides = useCollection("guides");
  const file = useRef<HTMLInputElement>(null);
  const [state, setState] = useState<{ busy?: string; done?: string; error?: string }>({});
  const run = (busy: string, task: () => Promise<string>) => (setState({ busy }), task().then((done) => setState({ done }), (e) => setState({ error: e instanceof Error ? e.message : String(e) })));
  return (
    <Section
      title="Guides"
      description={`The guide of each visual of the reports (Trainer), kept in this browser: ${guides.length} guide(s). The Excel / CSV file has the columns id, report_id, page, visual, title, content.`}
      actions={
        <>
          <input ref={file} type="file" accept=".xlsx,.csv" hidden onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = "";
            if (f) run("Importing…", async () => `${await importGuides(f)} guide(s) imported.`);
          }} />
          <Button size="sm" variant="outline" onClick={() => file.current?.click()}>
            <UploadIcon /> Import Excel / CSV
          </Button>
          <Button size="sm" variant="outline" disabled={!guides.length} onClick={() => run("Writing the file…", async () => (await downloadGuides(), "Downloaded."))}>
            <DownloadIcon /> Download as Excel
          </Button>
        </>
      }
    >
      {state.busy && <p className="text-[13px] text-muted-foreground">{state.busy}</p>}
      {state.done && <p className="text-[13px] text-success">{state.done}</p>}
      {state.error && <p className="text-[13px] text-destructive">{state.error}</p>}
    </Section>
  );
}

