import { useRef } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { BracesIcon, PlusIcon, UploadIcon } from "lucide-react";
import { EmptyState, Page } from "@/components/page";
import { Button } from "@/components/ui/button";
import { newSchema, schemaSchema } from "@/lib/schema";
import { schemas } from "@/lib/store";

export const Route = createFileRoute("/schemas/")({ component: Schemas });

function Schemas() {
  const all = schemas.use();
  const navigate = useNavigate();
  const input = useRef<HTMLInputElement>(null);

  const create = () => {
    const schema = newSchema(`Schema ${all.length + 1}`);
    schemas.put(schema);
    void navigate({ to: "/schemas/$id", params: { id: schema.id } });
  };

  const importFile = async (file: File) => {
    try {
      const schema = schemaSchema.parse({ ...JSON.parse(await file.text()), id: crypto.randomUUID() });
      schemas.put(schema);
      void navigate({ to: "/schemas/$id", params: { id: schema.id } });
    } catch (error) {
      alert(`Invalid schema file: ${error instanceof Error ? error.message : error}`);
    }
  };

  return (
    <Page
      title="Schemas"
      subtitle="Structure of the rows of a file: fields, types, options"
      actions={
        <div className="flex gap-1.5">
          <input ref={input} type="file" accept=".json" hidden onChange={(e) => e.target.files?.[0] && void importFile(e.target.files[0])} />
          <Button size="sm" variant="ghost" onClick={() => input.current?.click()}>
            <UploadIcon /> Import
          </Button>
          <Button size="sm" variant="outline" onClick={create}>
            <PlusIcon /> New schema
          </Button>
        </div>
      }
    >
      {all.length === 0 ? (
        <EmptyState icon={<BracesIcon />} title="No schema yet">
          A schema lists the columns of a file and their types. The agent can help you write it.
        </EmptyState>
      ) : (
        <ul className="divide-y overflow-hidden rounded-xl border bg-card shadow-soft">
          {all.map((schema) => (
            <li key={schema.id} className="relative flex items-center gap-3 px-3.5 py-2.5 transition-colors hover:bg-muted/50">
              <BracesIcon className="size-4 shrink-0 text-muted-foreground" />
              <Link to="/schemas/$id" params={{ id: schema.id }} className="min-w-0 flex-1 after:absolute after:inset-0">
                <span className="block truncate font-medium">{schema.name}</span>
                <span className="block truncate text-[12px] text-muted-foreground">
                  {schema.fields.length} fields · sheet “{schema.sheet}”
                </span>
              </Link>
            </li>
          ))}
        </ul>
      )}
    </Page>
  );
}
