import { useState } from "react";
import { FileUpIcon } from "lucide-react";
import { AddDialog } from "./entity";
import { createDatasetFromCsv } from "@/lib/library";
import { refreshModel } from "@/server/ms";

/** A CSV file becomes a semantic model (file kept in OneDrive › Hub/Datasets), refreshed once to load it. */
export function CsvDialog({ open, onOpenChange, onCreated }: { open: boolean; onOpenChange: (v: boolean) => void; onCreated: (id: string) => void }) {
  const [file, setFile] = useState<File>();
  return (
    <AddDialog
      open={open}
      onOpenChange={onOpenChange}
      title="New dataset from a CSV file"
      description="The file is put in your OneDrive (Hub/Datasets) and a semantic model reads it; it is then refreshed once to load the data."
      fields={[{ name: "name", label: "Name of the dataset" }]}
      submit="Upload and create"
      extra={
        <label className="flex cursor-pointer items-center gap-3 rounded-lg border border-dashed p-4 text-[13px] hover:bg-accent">
          <FileUpIcon className="size-5 text-muted-foreground" />
          <span className="flex-1 truncate">{file ? file.name : "Choose a .csv file"}</span>
          <input type="file" accept=".csv,text/csv" className="sr-only" onChange={(e) => setFile(e.target.files?.[0])} />
        </label>
      }
      onSubmit={async ({ name }) => {
        if (!file) throw new Error("Choose a CSV file.");
        const entry = await createDatasetFromCsv(name, file);
        await refreshModel({ data: { datasetId: entry.id, groupId: entry.groupId } }).catch(() => undefined);
        onCreated(entry.id);
      }}
    />
  );
}

