/** Pickers of the screens' items, each with its "Add" dialog (reports, datasets). */
import { useState } from "react";
import { AddDialog } from "./entity";
import { EntityPicker } from "./workspace";
import { addDataset, addReport, createReportFromModel, notEditable } from "@/lib/library";
import { useCollection } from "@/lib/store";

/**
 * Reports known to the app. `editable` (Builder): only those whose definition can be changed; a report
 * added that is not editable is refused with the reason; a new report can be made from a model.
 */
export function ReportPicker({ value, onChange, editable = false }: { value?: string; onChange: (id: string) => void; editable?: boolean }) {
  const reports = useCollection("reports").filter((r) => !editable || r.access === "edit");
  const [adding, setAdding] = useState<"link" | "model">();
  return (
    <>
      <EntityPicker
        items={reports.map((r) => ({ id: r.id, label: r.name, hint: r.access === "view" ? "view only" : undefined }))}
        value={value}
        onChange={onChange}
        placeholder={reports.length ? "Choose a report" : "Add a report"}
        onAdd={() => setAdding("link")}
        addLabel="Add a report…"
        more={editable ? [{ label: "New report from a semantic model…", onSelect: () => setAdding("model") }] : undefined}
      />
      <AddDialog
        open={adding === "link"}
        onOpenChange={(o) => setAdding(o ? "link" : undefined)}
        title="Add a report"
        description={editable ? "Paste the link of a Power BI report saved in PBIR format (editable). A report made in the Power BI service is often in the older format: create a new report from its semantic model instead." : "Paste the link of a Power BI report (workspace or app). Its semantic model is added with it."}
        fields={[{ name: "link", label: "Report link", placeholder: "https://app.powerbi.com/groups/…/reports/…" }]}
        onSubmit={async ({ link }) => {
          const entry = await addReport(link);
          // Added to the app either way, but the Builder only opens what it can edit: say why.
          if (editable && entry.access !== "edit") throw new Error(notEditable(entry));
          onChange(entry.id);
        }}
      />
      {editable && (
        <AddDialog
          open={adding === "model"}
          onOpenChange={(o) => setAdding(o ? "model" : undefined)}
          title="New report from a semantic model"
          description="A new report, live-connected to the model, is created in My workspace with a blank page; the assistant then builds its pages. Needs the Build permission on the model."
          fields={[
            { name: "model", label: "Semantic model link", placeholder: "https://app.powerbi.com/groups/…/datasets/…" },
            { name: "name", label: "Report name", optional: true },
          ]}
          submit="Create"
          onSubmit={async ({ model, name }) => onChange((await createReportFromModel(model, name)).id)}
        />
      )}
    </>
  );
}

/** Semantic models known to the app (`editable`: only those whose definition can be changed). */
export function DatasetPicker({ value, onChange, editable = false }: { value?: string; onChange: (id: string) => void; editable?: boolean }) {
  const datasets = useCollection("datasets").filter((d) => !editable || d.access === "edit");
  const [adding, setAdding] = useState(false);
  return (
    <>
      <EntityPicker
        items={datasets.map((d) => ({ id: d.id, label: d.name, hint: d.access === "edit" ? undefined : d.access }))}
        value={value}
        onChange={onChange}
        placeholder={datasets.length ? "Choose a dataset" : "Add a dataset"}
        onAdd={() => setAdding(true)}
        addLabel="Add a dataset…"
      />
      <AddDialog
        open={adding}
        onOpenChange={setAdding}
        title="Add a dataset"
        description="Paste the link of a Power BI semantic model (its settings or details page)."
        fields={[{ name: "link", label: "Semantic model link", placeholder: "https://app.powerbi.com/groups/…/datasets/…" }]}
        onSubmit={async ({ link }) => onChange((await addDataset(link)).id)}
      />
    </>
  );
}
