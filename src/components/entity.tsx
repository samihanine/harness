/**
 * Adding and managing the items of a screen (reports, datasets, Excel files): there is no library
 * anymore, each screen adds its own items by link and offers rename / read again / download / import /
 * remove in a "…" menu.
 */
import { useRef, useState, type ReactNode } from "react";
import { DownloadIcon, EllipsisIcon, FileUpIcon, PencilIcon, RefreshCwIcon, Trash2Icon } from "lucide-react";
import { toast } from "sonner";
import { Button } from "./ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "./ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { Input } from "./ui/input";
import { Label } from "./ui/label";
import { downloadEntry, importEntries, type Kind } from "@/lib/library";
import { errorText } from "@/services/ms";

export type FieldSpec = { name: string; label: string; placeholder?: string; hint?: string; optional?: boolean };

/** A dialog asking for a few values (links, names) then running `onSubmit`; errors stay shown in it. */
export function AddDialog({ open, onOpenChange, title, description, fields, submit = "Add", onSubmit, extra }: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  title: string;
  description?: ReactNode;
  fields: FieldSpec[];
  submit?: string;
  onSubmit: (values: Record<string, string>) => Promise<unknown>;
  extra?: ReactNode;
}) {
  const [values, setValues] = useState<Record<string, string>>({});
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const ready = fields.every((f) => f.optional || values[f.name]?.trim());
  const run = () => {
    setState({ busy: true });
    onSubmit(Object.fromEntries(fields.map((f) => [f.name, values[f.name]?.trim() ?? ""]))).then(
      () => {
        setState({});
        setValues({});
        onOpenChange(false);
      },
      (e) => setState({ error: errorText(e) }),
    );
  };
  return (
    <Dialog open={open} onOpenChange={(v) => (setState({}), onOpenChange(v))}>
      <DialogContent className="sm:max-w-lg">
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
          {description && <DialogDescription>{description}</DialogDescription>}
        </DialogHeader>
        <form
          className="flex flex-col gap-4"
          onSubmit={(e) => {
            e.preventDefault();
            if (ready && !state.busy) run();
          }}
        >
          {fields.map((f, i) => (
            <div key={f.name} className="flex flex-col gap-1.5">
              <Label htmlFor={`add-${f.name}`}>
                {f.label}
                {f.optional && <span className="font-normal text-muted-foreground"> (optional)</span>}
              </Label>
              <Input id={`add-${f.name}`} autoFocus={i === 0} placeholder={f.placeholder} value={values[f.name] ?? ""} onChange={(e) => setValues({ ...values, [f.name]: e.target.value })} />
              {f.hint && <p className="text-[12px] text-muted-foreground">{f.hint}</p>}
            </div>
          ))}
          {extra}
          {state.error && <p className="rounded-md bg-destructive/10 px-3 py-2 text-[12.5px] break-words text-destructive">{state.error}</p>}
          <DialogFooter>
            <Button type="button" variant="outline" onClick={() => onOpenChange(false)}>
              Cancel
            </Button>
            <Button type="submit" disabled={!ready || state.busy}>
              {state.busy ? "Working…" : submit}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

/** Runs a task with a toast: loading, then done or the error. */
export function withToast<T>(label: string, task: Promise<T>, done?: string) {
  toast.promise(task, { loading: label, success: done ?? "Done", error: (e) => errorText(e) });
  return task;
}

/** The "…" menu of the chosen item. */
export function EntityMenu({ kind, entry, onRename, onRefresh, onRemove, children }: {
  kind: Kind;
  entry?: { id: string; name?: string; title?: string };
  onRename?: () => void;
  onRefresh?: () => void;
  onRemove?: () => void;
  /** More items (screen specific), on top. */
  children?: ReactNode;
}) {
  const input = useRef<HTMLInputElement>(null);
  return (
    <>
      <input
        ref={input}
        type="file"
        multiple
        accept=".json,.md,.markdown,.txt"
        hidden
        onChange={(e) => {
          const files = e.target.files;
          if (files?.length) void withToast("Importing…", importEntries(files), "Imported.");
          e.target.value = "";
        }}
      />
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <Button variant="outline" size="icon" className="size-9" aria-label="More actions">
            <EllipsisIcon />
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="end" className="w-56">
          {children}
          {children && <DropdownMenuSeparator />}
          {entry && onRename && (
            <DropdownMenuItem onSelect={onRename}>
              <PencilIcon /> Rename (in this app)
            </DropdownMenuItem>
          )}
          {entry && onRefresh && (
            <DropdownMenuItem onSelect={onRefresh}>
              <RefreshCwIcon /> Read again
            </DropdownMenuItem>
          )}
          {entry && (
            <DropdownMenuItem onSelect={() => downloadEntry(kind, entry)}>
              <DownloadIcon /> Download (JSON)
            </DropdownMenuItem>
          )}
          <DropdownMenuItem onSelect={() => input.current?.click()}>
            <FileUpIcon /> Import (JSON)
          </DropdownMenuItem>
          {entry && onRemove && (
            <>
              <DropdownMenuSeparator />
              <DropdownMenuItem variant="destructive" onSelect={onRemove}>
                <Trash2Icon /> Remove from the app
              </DropdownMenuItem>
            </>
          )}
        </DropdownMenuContent>
      </DropdownMenu>
    </>
  );
}
