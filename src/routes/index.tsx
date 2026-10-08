import { useState, type ReactNode } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { CopyIcon, ExternalLinkIcon, PlusIcon, RefreshCwIcon, Trash2Icon } from "lucide-react";
import { resolveFolder, imagesFolderLabel } from "@/lib/excel";
import { addDataset, addExcel, addReport, cloneReport, refreshDatasetText } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { patch, remove, useCollection, type ExcelEntry } from "@/lib/store";

export const Route = createFileRoute("/")({ component: Library });

function Library() {
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const excels = useCollection("excels");
  return (
    <div className="grid h-full grid-cols-3 gap-3 overflow-auto p-3">
      <Section title="Reports" placeholder="Power BI report link" onAdd={addReport}>
        {reports.map((r) => (
          <Item
            key={r.id}
            title={r.name}
            subtitle={`${r.editable === false ? "view only" : r.editable ? "editable" : ""} · model ${datasets.find((d) => d.id === r.datasetId)?.name ?? r.datasetId}`}
            href={r.url}
            context={r.context}
            onContext={(context) => patch("reports", r.id, { context })}
            onRefresh={() => addReport(r.url)}
            onRemove={() => remove("reports", r.id)}
            extra={
              <>
                <div className="grid grid-cols-2 gap-2">
                  <Pick label="Info links (Excel)" value={r.infoExcelId} options={excels} onChange={(infoExcelId) => patch("reports", r.id, { infoExcelId })} />
                  <Pick label="Guides (Excel)" value={r.guidesExcelId} options={excels} onChange={(guidesExcelId) => patch("reports", r.id, { guidesExcelId })} />
                </div>
                <Action icon={<CopyIcon />} label="Clone to My workspace" run={() => cloneReport(r)} />
                {r.snapshot && (
                  <details>
                    <summary className="label cursor-pointer">What the AI knows of it</summary>
                    <pre className="mt-1 max-h-60 overflow-auto rounded bg-muted p-2 text-[11px] whitespace-pre-wrap">{r.snapshot}</pre>
                  </details>
                )}
              </>
            }
          />
        ))}
      </Section>
      <Section title="Semantic models" placeholder="Semantic model link or id" onAdd={(v) => addDataset(v)}>
        {datasets.map((d) => (
          <Item
            key={d.id}
            title={d.name}
            subtitle={d.excelId ? `from Excel ${excels.find((e) => e.id === d.excelId)?.name ?? ""}` : d.groupId ? "workspace" : "My workspace"}
            href={`https://app.powerbi.com/${d.groupId ? `groups/${d.groupId}` : "groups/me"}/datasets/${d.id}/details`}
            context={d.context}
            onContext={(context) => patch("datasets", d.id, { context })}
            onRefresh={() => refreshDatasetText(d)}
            onRemove={() => remove("datasets", d.id)}
            extra={
              d.model && (
                <details>
                  <summary className="label cursor-pointer">Structure</summary>
                  <pre className="mt-1 max-h-60 overflow-auto rounded bg-muted p-2 text-[11px] whitespace-pre-wrap">{d.model}</pre>
                </details>
              )
            }
          />
        ))}
      </Section>
      <Section title="Excel files (SharePoint / OneDrive)" placeholder="Sharing link of an .xlsx" onAdd={(v) => addExcel(v)}>
        {excels.map((e) => (
          <Item
            key={e.id}
            title={e.name}
            subtitle={`table ${e.table}`}
            href={e.link}
            context={e.context}
            onContext={(context) => patch("excels", e.id, { context })}
            onRemove={() => remove("excels", e.id)}
            extra={<ImagesFolder excel={e} />}
          />
        ))}
      </Section>
    </div>
  );
}

function Section({ title, placeholder, onAdd, children }: { title: string; placeholder: string; onAdd: (value: string) => Promise<unknown>; children: ReactNode }) {
  const [value, setValue] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const add = async () => {
    setState({ busy: true });
    try {
      await onAdd(value.trim());
      setValue("");
      setState({});
    } catch (e) {
      setState({ error: errorText(e) });
    }
  };
  return (
    <section className="flex min-h-0 flex-col gap-2">
      <h2 className="font-semibold">{title}</h2>
      <div className="flex gap-1">
        <input className="input" placeholder={placeholder} value={value} onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => e.key === "Enter" && value && void add()} />
        <button type="button" className="btn" disabled={!value || state.busy} onClick={() => void add()}>
          <PlusIcon /> {state.busy ? "Reading…" : "Add"}
        </button>
      </div>
      {state.error && <p className="text-[12px] text-destructive">{state.error}</p>}
      <div className="flex flex-col gap-2">{children}</div>
    </section>
  );
}

function Item(props: {
  title: string;
  subtitle?: string;
  href?: string;
  context: string;
  onContext: (value: string) => void;
  onRefresh?: () => Promise<unknown>;
  onRemove: () => void;
  extra?: ReactNode;
}) {
  return (
    <details className="group rounded-lg border bg-card">
      <summary className="flex cursor-pointer list-none items-center gap-2 px-3 py-2">
        <div className="min-w-0 flex-1">
          <div className="truncate font-medium">{props.title}</div>
          {props.subtitle && <div className="truncate text-[11px] text-muted-foreground">{props.subtitle}</div>}
        </div>
        {props.href && (
          <a className="icon-btn" href={props.href} target="_blank" rel="noreferrer" title="Open" onClick={(e) => e.stopPropagation()}>
            <ExternalLinkIcon />
          </a>
        )}
      </summary>
      <div className="flex flex-col gap-2 border-t p-3">
        <label className="label">Context for the AI</label>
        <textarea className="input h-20 py-1" defaultValue={props.context} onBlur={(e) => e.target.value !== props.context && props.onContext(e.target.value)} />
        {props.extra}
        <div className="flex gap-1">
          {props.onRefresh && <Action icon={<RefreshCwIcon />} label="Read again" run={props.onRefresh} />}
          <button type="button" className="btn ml-auto text-destructive" onClick={props.onRemove}>
            <Trash2Icon /> Remove
          </button>
        </div>
      </div>
    </details>
  );
}

function Action({ icon, label, run }: { icon: ReactNode; label: string; run: () => Promise<unknown> }) {
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  return (
    <span className="flex flex-col gap-1">
      <button
        type="button"
        className="btn w-fit"
        disabled={state.busy}
        onClick={() => {
          setState({ busy: true });
          run().then(() => setState({}), (e) => setState({ error: errorText(e) }));
        }}
      >
        {icon} {state.busy ? "…" : label}
      </button>
      {state.error && <span className="text-[11px] text-destructive">{state.error}</span>}
    </span>
  );
}

function Pick({ label, value, options, onChange }: { label: string; value?: string; options: { id: string; name: string }[]; onChange: (id?: string) => void }) {
  return (
    <label className="flex flex-col gap-1">
      <span className="label">{label}</span>
      <select className="input" value={value ?? ""} onChange={(e) => onChange(e.target.value || undefined)}>
        <option value="">None</option>
        {options.map((o) => (
          <option key={o.id} value={o.id}>
            {o.name}
          </option>
        ))}
      </select>
    </label>
  );
}

/** Where uploaded images go: a SharePoint / OneDrive folder link, or the default folder next to the file. */
function ImagesFolder({ excel }: { excel: ExcelEntry }) {
  const [error, setError] = useState<string>();
  return (
    <label className="flex flex-col gap-1">
      <span className="label">Images folder (sharing link of a folder; default: {imagesFolderLabel({ ...excel, imagesFolder: undefined })})</span>
      <input
        className="input"
        defaultValue={excel.imagesFolder?.link ?? ""}
        placeholder="https://…sharepoint.com/:f:/…"
        onBlur={(e) => {
          const link = e.target.value.trim();
          if (link === (excel.imagesFolder?.link ?? "")) return;
          setError(undefined);
          if (!link) return void patch("excels", excel.id, { imagesFolder: undefined });
          resolveFolder(link).then((imagesFolder) => patch("excels", excel.id, { imagesFolder }), (err) => setError(errorText(err)));
        }}
      />
      {excel.imagesFolder && <span className="text-[11px] text-muted-foreground">Uploads go to "{excel.imagesFolder.name}".</span>}
      {error && <span className="text-[11px] text-destructive">{error}</span>}
    </label>
  );
}
