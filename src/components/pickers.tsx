/** Pickers of the screens' items, each with its "Add" dialog (reports, datasets). */
import { useState } from "react";
import { AddDialog } from "./entity";
import { EntityPicker } from "./workspace";
import { addDataset, addReport } from "@/lib/library";
import { useCollection } from "@/lib/store";

/** Reports known to the app (`editable`: only those whose definition can be changed). */
export function ReportPicker({ value, onChange, editable = false }: { value?: string; onChange: (id: string) => void; editable?: boolean }) {
  const reports = useCollection("reports").filter((r) => !editable || r.access === "edit");
  const [adding, setAdding] = useState(false);
  return (
    <>
      <EntityPicker
        items={reports.map((r) => ({ id: r.id, label: r.name, hint: r.access === "view" ? "view only" : undefined }))}
        value={value}
        onChange={onChange}
        placeholder={reports.length ? "Choose a report" : "Add a report"}
        onAdd={() => setAdding(true)}
        addLabel="Add a report…"
      />
      <AddDialog
        open={adding}
        onOpenChange={setAdding}
        title="Add a report"
        description="Paste the link of a Power BI report (workspace or app). Its semantic model is added with it."
        fields={[{ name: "link", label: "Report link", placeholder: "https://app.powerbi.com/groups/…/reports/…" }]}
        onSubmit={async ({ link }) => onChange((await addReport(link)).id)}
      />
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
