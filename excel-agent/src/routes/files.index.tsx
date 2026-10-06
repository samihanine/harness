import { useState } from "react";
import { Link, createFileRoute } from "@tanstack/react-router";
import { FileSpreadsheetIcon, PlusIcon, Trash2Icon } from "lucide-react";
import { IconButton } from "@/components/icon-button";
import { NewFileDialog } from "@/components/new-file-dialog";
import { EmptyState, Page, ago } from "@/components/page";
import { Button } from "@/components/ui/button";
import { handles, links, schemas } from "@/lib/store";

export const Route = createFileRoute("/files/")({ component: Files });

function Files() {
  const files = links.use();
  const allSchemas = schemas.use();
  const [creating, setCreating] = useState(false);
  const sorted = [...files].sort((a, b) => b.openedAt - a.openedAt);

  return (
    <Page
      title="Files"
      subtitle="Local Excel files edited with a schema"
      actions={
        <Button size="sm" variant="outline" onClick={() => setCreating(true)}>
          <PlusIcon /> Add file
        </Button>
      }
    >
      {sorted.length === 0 ? (
        <EmptyState icon={<FileSpreadsheetIcon />} title="No file yet">
          Add an Excel file from your computer and pick the schema describing its rows.
        </EmptyState>
      ) : (
        <ul className="divide-y overflow-hidden rounded-xl border bg-card shadow-soft">
          {sorted.map((file) => (
            <li key={file.id} className="group relative flex items-center gap-3 px-3.5 py-2.5 transition-colors hover:bg-muted/50">
              <FileSpreadsheetIcon className="size-4 shrink-0 text-muted-foreground" />
              <Link to="/files/$id" params={{ id: file.id }} className="min-w-0 flex-1 after:absolute after:inset-0">
                <span className="block truncate font-medium">{file.name}</span>
                <span className="block truncate text-[12px] text-muted-foreground">
                  {file.fileName} · {allSchemas.find((s) => s.id === file.schemaId)?.name ?? "missing schema"}
                </span>
              </Link>
              <span className="text-[12px] text-muted-foreground group-hover:hidden">{ago(file.openedAt)}</span>
              <IconButton
                label="Remove from the list"
                className="relative hidden group-hover:inline-flex"
                onClick={() => {
                  links.remove(file.id);
                  void handles.remove(file.id);
                }}
              >
                <Trash2Icon />
              </IconButton>
            </li>
          ))}
        </ul>
      )}
      <NewFileDialog open={creating} onClose={() => setCreating(false)} />
    </Page>
  );
}
