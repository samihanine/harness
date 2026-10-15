import { useEffect, useRef, useState } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { DatabaseIcon, ExternalLinkIcon, PlusIcon, RefreshCwIcon, SparklesIcon, Trash2Icon, XIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { Picker } from "@/components/picker";
import { Popover } from "@/components/popover";
import { Empty, PageHeader, SidePanel } from "@/components/side-panel";
import { AI_PANEL } from "@/config";
import { withContexts } from "@/agent/context";
import type { Agent } from "@/agent/loop";
import { DAX_RULES } from "@/agent/pbi-tools";
import { datasetsContext, researchState, researchTools } from "@/agent/research-tools";
import { previewUrl } from "@/lib/excel";
import { useLast } from "@/lib/last";
import { errorText } from "@/lib/ms";
import { publishResearch, researchFile } from "@/lib/research";
import { Drawer } from "@/components/drawer";
import { find, patch, put, remove, useCollection, type ResearchEntry } from "@/lib/store";

export const Route = createFileRoute("/research")({
  validateSearch: (s: Record<string, unknown>) => ({ research: (s.research as string) || undefined }),
  component: Researcher,
});

/**
 * Researcher: analyses across several semantic models, built by the AI as sheets of an Excel workbook
 * (PivotTables connected to the models, formula tables, charts) — Excel computes everything. The
 * workbook lives on the user's OneDrive (Hub/Researcher) and is shown read only.
 */
function Researcher() {
  const { research: researchId } = Route.useSearch();
  const navigate = useNavigate({ from: "/research" });
  useLast("research:research", researchId, (research) => navigate({ search: { research }, replace: true }));
  const researches = useCollection("researches").sort((a, b) => b.updatedAt - a.updatedAt);
  const allDatasets = useCollection("datasets").filter((d) => d.access !== "view");
  const research = researches.find((r) => r.id === researchId);
  const datasets = allDatasets.filter((d) => research?.datasetIds.includes(d.id));
  const [state, setState] = useState<{ busy?: string; error?: string }>({});
  /** Bumped after each publish: the preview is loaded again. */
  const [version, setVersion] = useState(0);
  const live = useRef({ research, datasets, dirty: false });
  live.current = { ...live.current, research, datasets };

  const publish = async () => {
    const r = live.current.research && (await find("researches", live.current.research.id));
    if (!r) return [];
    setState({ busy: "Publishing…" });
    try {
      const { errors } = await publishResearch(r, (step) => setState({ busy: `${step}…` }));
      live.current.dirty = false;
      setVersion((v) => v + 1);
      setState(errors.length ? { error: errors.join(" · ") } : {});
      return errors;
    } catch (e) {
      setState({ error: errorText(e) });
      throw e;
    }
  };
  const [creating, setCreating] = useState(false);
  /** A research works in an Excel file given by link: its content is replaced by the research's sheets. */
  const create = async (link: string) => {
    const { name, file } = await researchFile(link);
    const r = await put("researches", { id: crypto.randomUUID(), name, datasetIds: [], sheets: [], file, updatedAt: Date.now() });
    setCreating(false);
    navigate({ search: { research: r.id } });
  };

  const agent: Agent | undefined = research && {
    instructions: `You are a data analyst. You build analyses in an Excel workbook from the user's semantic models, one analysis per sheet:
- create_table_from_tcd: a real Excel PivotTable connected to a model (rows / columns / filters = columns, values = measures of the model);
- create_table + create_column: an Excel table whose columns are formulas (they may refer to other tables of the workbook);
- create_chart: a chart of a sheet.
You NEVER type data yourself: values come from the models (pivots) or from formulas. Use run_dax and get_table_values to think (look at values, check results), run_js to compute on values. Do the whole request in one go, then answer with what you built and the key findings (read them with get_table_values).
${DAX_RULES}`,
    tools: researchTools({
      research: () => live.current.research,
      datasets: () => live.current.datasets,
      setSheets: async (sheets) => {
        const r = live.current.research!;
        live.current.research = { ...r, sheets };
        live.current.dirty = true;
        await patch("researches", r.id, { sheets, updatedAt: Date.now() });
      },
      publish: async () => {
        if (live.current.dirty || !live.current.research?.file) await publish();
      },
    }),
    context: () => withContexts([`Research "${research.name}".`, datasetsContext(datasets)], datasets.map((d) => d.contextIds)),
    state: async () => researchState(live.current.research, live.current.datasets, live.current.dirty),
    finish: async () => {
      if (!live.current.dirty) return undefined;
      const errors = await publish();
      return errors.length ? `⚠️ Published, with errors: ${errors.join(" · ")}` : "✓ Workbook published.";
    },
  };

  return (
    <div className="flex h-full">
      {AI_PANEL && (
        <SidePanel
          tabs={[
            {
              id: "ai",
              label: "AI assistant",
              icon: SparklesIcon,
              content: !research ? <Empty>Choose or create a research.</Empty> : !datasets.length ? <Empty>Add the semantic models of this research (above the workbook) first.</Empty> : <Chat scope={`research:${research.id}`} agent={agent!} placeholder="Ask for an analysis: a pivot, a comparison, a chart…" />,
            },
          ]}
        />
      )}
      <div className="flex min-w-0 flex-1 flex-col">
        <PageHeader
          screen="Researcher"
          picker={
            <Picker
              items={researches.map((r) => ({ id: r.id, label: r.name, hint: `${r.sheets.length} sheet${r.sheets.length === 1 ? "" : "s"}` }))}
              value={researchId}
              onChange={(research) => navigate({ search: { research } })}
              placeholder="Choose a research…"
              empty="No research yet."
              footer={
                <button type="button" className="flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-left hover:bg-muted" onClick={() => setCreating(true)}>
                  <PlusIcon className="size-4" /> New research
                </button>
              }
            />
          }
          actions={
            <>
              {!research && (
                <button type="button" className="btn-primary" onClick={() => setCreating(true)}>
                  <PlusIcon /> New research
                </button>
              )}
              {research && (
                <>
                  {state.busy && <span className="text-[13px] text-muted-foreground">{state.busy}</span>}
                  <button type="button" className="btn" disabled={!!state.busy} title="Write the workbook again and refresh its pivots from the models" onClick={() => void publish().catch(() => undefined)}>
                    <RefreshCwIcon className={state.busy ? "animate-spin" : ""} /> Refresh
                  </button>
                  {research.file && (
                    <a className="btn" href={research.file.webUrl} target="_blank" rel="noreferrer">
                      <ExternalLinkIcon /> Open in Excel
                    </a>
                  )}
                  <button type="button" className="icon-btn hover:text-destructive" title="Delete the research (its workbook stays on OneDrive)" onClick={() => confirm(`Delete the research "${research.name}"? (its workbook stays on OneDrive)`) && void remove("researches", research.id).then(() => navigate({ search: { research: undefined } }))}>
                    <Trash2Icon />
                  </button>
                </>
              )}
            </>
          }
          below={research && <Datasets research={research} />}
        >
          {state.error && <span className="max-w-xl truncate text-[13px] text-destructive" title={state.error}>{state.error}</span>}
        </PageHeader>
        {!research && <Empty>A research is an Excel workbook (given by link) where the AI builds analyses from your semantic models: real PivotTables connected to the models, formula tables, charts.</Empty>}
        {research && <Preview key={`${research.id}:${version}`} research={research} />}
      </div>
      {creating && <NewResearch onCreate={create} onClose={() => setCreating(false)} />}
    </div>
  );
}

/** A new research: the link of the Excel file it works in (an empty workbook is best: its content is replaced). */
function NewResearch({ onCreate, onClose }: { onCreate: (link: string) => Promise<void>; onClose: () => void }) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  return (
    <Drawer
      title="New research"
      onClose={onClose}
      footer={
        <>
          <button
            type="button"
            className="btn-primary"
            disabled={!link.trim() || state.busy}
            onClick={() => {
              setState({ busy: true });
              onCreate(link.trim()).catch((e) => setState({ error: errorText(e) }));
            }}
          >
            {state.busy ? "Reading…" : "Create"}
          </button>
          <button type="button" className="btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <label className="flex flex-col gap-1.5">
        <span className="label">Excel file of the research</span>
        <input autoFocus className="input" placeholder="Sharing link of an .xlsx (SharePoint / OneDrive)" value={link} onChange={(e) => setLink(e.target.value)} />
        <span className="text-[12px] text-muted-foreground">The research writes its sheets (pivots, formula tables, charts) into this workbook: its current content is replaced. Use a new, empty workbook.</span>
      </label>
      {state.error && <p className="text-[13px] break-words text-destructive">{state.error}</p>}
    </Drawer>
  );
}

/** The models of the research (chips), added from the library. */
function Datasets({ research }: { research: ResearchEntry }) {
  const all = useCollection("datasets").filter((d) => d.access !== "view");
  const set = (datasetIds: string[]) => void patch("researches", research.id, { datasetIds });
  return (
    <>
      <span className="text-[13.5px] text-muted-foreground">Models</span>
      {research.datasetIds.map((id) => {
        const d = all.find((x) => x.id === id);
        return (
          <span key={id} className="chip pr-1.5">
            <DatabaseIcon /> {d?.name ?? id}
            <button type="button" className="icon-btn size-6" title="Remove from the research" onClick={() => set(research.datasetIds.filter((x) => x !== id))}>
              <XIcon />
            </button>
          </span>
        );
      })}
      <Popover
        width={320}
        trigger={(open) => (
          <button type="button" className="chip" onClick={open}>
            <PlusIcon /> Add a model
          </button>
        )}
      >
        {(close) => (
          <div className="flex max-h-72 flex-col overflow-y-auto">
            {all.filter((d) => !research.datasetIds.includes(d.id)).map((d) => (
              <button key={d.id} type="button" className="flex items-center gap-2 rounded-lg px-2.5 py-2 text-left hover:bg-muted" onClick={() => (set([...research.datasetIds, d.id]), close())}>
                <DatabaseIcon className="size-4 text-muted-foreground" /> {d.name}
              </button>
            ))}
            {all.length === 0 && (
              <p className="p-2 text-[13px] text-muted-foreground">
                No model with Build rights in the <Link to="/" search={{ tab: "datasets" }} className="underline">Library</Link>.
              </p>
            )}
          </div>
        )}
      </Popover>
    </>
  );
}

/** The workbook as Excel for the web shows it (read only); the link opening it expires after a few minutes, so it is asked again. */
function Preview({ research }: { research: ResearchEntry }) {
  const [url, setUrl] = useState<string>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    if (research.file) previewUrl(research.file).then(setUrl, (e) => setError(errorText(e)));
  }, [research.file?.itemId]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!research.file) return <Empty>The workbook is created with the first analysis (ask the AI), or with Refresh.</Empty>;
  return (
    <div className="flex min-h-0 flex-1 flex-col p-6 pt-4">
      {error && <p className="text-[13px] text-destructive">{error}</p>}
      <div className="card min-h-0 flex-1 overflow-hidden">{url ? <iframe title="Workbook" src={url} className="size-full border-0" /> : !error && <Empty>Loading the workbook…</Empty>}</div>
    </div>
  );
}
