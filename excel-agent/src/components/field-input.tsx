import { useEffect, useRef, useState } from "react";
import { BoldIcon, CheckIcon, ImageIcon, ItalicIcon, RemoveFormattingIcon, SearchIcon, UploadIcon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { TINTS } from "@/lib/colors";
import { normalizeRich } from "@/lib/rich";
import type { Field } from "@/lib/schema";
import { cn } from "@/lib/utils";
import { resolveImage } from "@/storage/images";
import type { ImageStore } from "@/storage/types";

export type InputProps = {
  field: Field;
  value: unknown;
  onChange: (value: unknown) => void;
  /** "cell": borderless, inside the table. "form": regular input. */
  variant: "cell" | "form";
  /** Where image fields store their files (local folder or SharePoint folder). */
  images?: ImageStore;
};

const base = {
  cell: "h-8 w-full bg-transparent px-2 outline-none focus:bg-background focus:ring-1 focus:ring-ring/60 rounded-sm",
  form: "h-8 w-full rounded-lg border bg-transparent px-2.5 outline-none focus:border-ring focus:ring-3 focus:ring-ring/30 transition-shadow",
};

/** Edits one value according to its field type (table cell or form). */
export function FieldInput(props: InputProps) {
  switch (props.field.type) {
    case "boolean":
      return (
        <div className={cn("flex items-center", props.variant === "cell" ? "h-8 px-2" : "h-8")}>
          <Checkbox checked={props.value === true} onCheckedChange={(checked) => props.onChange(checked)} />
        </div>
      );
    case "option":
      return <OptionInput {...props} />;
    case "image":
      return <ImageInput {...props} />;
    case "text":
      return props.variant === "form" ? <TextArea {...props} /> : <TextCell {...props} />;
    default:
      return <TextInput {...props} />;
  }
}

/** Text-like input committed on blur / Enter (so each keystroke does not write the file). */
function TextInput({ field, value, onChange, variant }: InputProps) {
  const shown = value === null || value === undefined ? "" : String(value);
  const [draft, setDraft] = useState(shown);
  useEffect(() => setDraft(shown), [shown]);
  const numeric = field.type === "number" || field.type === "integer";
  const commit = () => draft !== shown && onChange(draft);
  return (
    <input
      type={field.type === "date" ? "date" : "text"}
      inputMode={numeric ? "decimal" : undefined}
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") e.currentTarget.blur();
        if (e.key === "Escape") setDraft(shown);
      }}
      className={cn(base[variant], numeric && "text-right tabular-nums")}
    />
  );
}

const SWATCHES = ["#dc2626", "#ea580c", "#16a34a", "#2563eb", "#7c3aed", "#db2777"];

/** Rich text editor (bold, italic, colour); the value is canonical HTML, see lib/rich.ts. */
function TextArea({ value, onChange }: InputProps) {
  const shown = normalizeRich(value ?? "");
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (ref.current && ref.current.innerHTML !== shown) ref.current.innerHTML = shown;
  }, [shown]);
  const format = (command: string, arg?: string) => {
    ref.current?.focus();
    document.execCommand("styleWithCSS", false, "true");
    document.execCommand(command, false, arg);
  };
  const tool = "inline-flex size-6 items-center justify-center rounded hover:bg-muted";
  return (
    <div className="rounded-lg border transition-shadow focus-within:border-ring focus-within:ring-3 focus-within:ring-ring/30">
      <div className="flex items-center gap-0.5 border-b px-1.5 py-1" onMouseDown={(e) => e.preventDefault()}>
        <button type="button" title="Bold" className={tool} onClick={() => format("bold")}>
          <BoldIcon className="size-3.5" />
        </button>
        <button type="button" title="Italic" className={tool} onClick={() => format("italic")}>
          <ItalicIcon className="size-3.5" />
        </button>
        {SWATCHES.map((color) => (
          <button key={color} type="button" title="Colour" className={tool} onClick={() => format("foreColor", color)}>
            <span className="size-3 rounded-full" style={{ background: color }} />
          </button>
        ))}
        <button type="button" title="Clear formatting" className={tool} onClick={() => format("removeFormat")}>
          <RemoveFormattingIcon className="size-3.5" />
        </button>
      </div>
      <div
        ref={ref}
        contentEditable
        suppressContentEditableWarning
        onBlur={() => {
          const next = normalizeRich(ref.current?.innerHTML ?? "");
          if (next !== shown) onChange(next);
        }}
        className="field-sizing-content min-h-20 max-h-80 overflow-y-auto px-2.5 py-1.5 whitespace-pre-wrap outline-none"
      />
    </div>
  );
}

/** Long text in a cell: two lines, edited in a popover. */
function TextCell(props: InputProps) {
  return (
    <Popover>
      <PopoverTrigger className="flex min-h-8 w-full items-start px-2 py-1.5 text-left">
        <span className="line-clamp-2 whitespace-pre-wrap" dangerouslySetInnerHTML={{ __html: normalizeRich(props.value ?? "") }} />
      </PopoverTrigger>
      <PopoverContent align="start" className="w-96 p-2">
        <TextArea {...props} />
      </PopoverContent>
    </Popover>
  );
}

export function OptionBadges({ field, value }: { field: Field; value: unknown }) {
  const values = (Array.isArray(value) ? value : value ? [value] : []).map(String);
  return (
    <>
      {values.map((v) => {
        const option = field.options?.find((o) => o.value === v);
        return (
          <span
            key={v}
            className={cn(
              "inline-flex h-5 items-center rounded-full px-2 text-[12px] whitespace-nowrap",
              option ? TINTS[option.color] : "bg-destructive/10 text-destructive",
            )}
          >
            {v}
          </span>
        );
      })}
    </>
  );
}

function OptionInput({ field, value, onChange, variant }: InputProps) {
  const [search, setSearch] = useState("");
  const selected = (Array.isArray(value) ? value : value ? [value] : []).map(String);
  const options = (field.options ?? []).filter((o) => o.value.toLowerCase().includes(search.toLowerCase()));
  const toggle = (v: string) =>
    field.multiple
      ? onChange(selected.includes(v) ? selected.filter((s) => s !== v) : [...selected, v])
      : onChange(selected[0] === v ? null : v);
  return (
    <Popover onOpenChange={() => setSearch("")}>
      <PopoverTrigger
        className={cn(
          "flex min-h-8 w-full flex-wrap items-center gap-1 text-left",
          variant === "cell" ? "px-2 py-1" : "rounded-lg border px-2 py-1",
        )}
      >
        <OptionBadges field={field} value={value} />
        {selected.length === 0 && variant === "form" && <span className="text-muted-foreground">Select…</span>}
      </PopoverTrigger>
      <PopoverContent align="start" className="w-56 gap-0 p-1">
        {(field.options?.length ?? 0) > 6 && (
          <div className="flex items-center gap-1.5 border-b px-2 pb-1">
            <SearchIcon className="size-3.5 text-muted-foreground" />
            <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…" className="h-7 flex-1 bg-transparent outline-none" />
          </div>
        )}
        <div className="max-h-64 overflow-y-auto pt-1">
          {options.map((option) => (
            <button
              key={option.value}
              type="button"
              onClick={() => toggle(option.value)}
              className="flex w-full items-center gap-2 rounded-md px-2 py-1 text-left hover:bg-muted"
            >
              <span className={cn("inline-flex h-5 items-center rounded-full px-2 text-[12px]", TINTS[option.color])}>{option.value}</span>
              {selected.includes(option.value) && <CheckIcon className="ml-auto size-3.5" />}
            </button>
          ))}
          {options.length === 0 && <p className="px-2 py-1.5 text-[12px] text-muted-foreground">No option.</p>}
        </div>
      </PopoverContent>
    </Popover>
  );
}

/** Image of a cell value: a file of the images store, or a link (SharePoint, web). */
function Thumbnail({ store, value, className }: { store?: ImageStore; value: unknown; className?: string }) {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => {
    if (typeof value !== "string" || !value) return setUrl(null);
    void (/^https?:\/\//i.test(value) ? resolveImage(value) : (store?.resolve(value) ?? Promise.resolve(null))).then(setUrl);
  }, [store, value]);
  if (!value) return null;
  return url ? (
    <img src={url} alt="" className={cn("rounded object-cover ring-1 ring-foreground/10", className)} />
  ) : (
    <span className={cn("inline-flex items-center justify-center rounded bg-muted text-muted-foreground", className)} title={String(value)}>
      <ImageIcon className="size-3.5" />
    </span>
  );
}

function ImageInput({ value, onChange, variant, images }: InputProps) {
  const input = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<{ id: string; name: string }[] | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [link, setLink] = useState("");
  const run = async (action: () => Promise<string>) => {
    setBusy(true);
    setError("");
    try {
      onChange(await action());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <Popover
      onOpenChange={(open) => {
        if (!open || !images) return;
        setFiles(null);
        void images.list().then(setFiles, () => setFiles([]));
      }}
    >
      <PopoverTrigger className={cn("flex min-h-8 w-full items-center gap-2 text-left", variant === "cell" ? "px-2" : "rounded-lg border px-2 py-1")}>
        <Thumbnail store={images} value={value} className={variant === "cell" ? "size-6" : "size-12"} />
        {variant === "form" && <span className="truncate text-muted-foreground">{value ? String(value) : "No image"}</span>}
      </PopoverTrigger>
      <PopoverContent align="start" className="w-72 gap-2 p-2">
        {value ? <Thumbnail store={images} value={value} className="max-h-48 w-full" /> : null}
        <div className="flex gap-1.5">
          {images && (
            <>
              <input ref={input} type="file" accept="image/*" hidden onChange={(e) => e.target.files?.[0] && void run(() => images.save(e.target.files![0]))} />
              <Button size="sm" variant="outline" className="flex-1" disabled={busy} onClick={() => input.current?.click()}>
                <UploadIcon /> {busy ? "Uploading…" : `Upload to ${images.label}`}
              </Button>
            </>
          )}
          {!!value && (
            <Button size="sm" variant="ghost" onClick={() => onChange(null)}>
              <XIcon /> Clear
            </Button>
          )}
        </div>
        <input
          value={link}
          onChange={(e) => setLink(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && /^https?:\/\//i.test(link.trim())) {
              onChange(link.trim());
              setLink("");
            }
          }}
          placeholder="Or paste an image link (Enter)"
          className="h-7 rounded-md border bg-transparent px-2 text-[12px] outline-none focus:border-ring"
        />
        {!images && <p className="text-[11px] text-muted-foreground">To upload images, choose an images folder with the folder button in the toolbar.</p>}
        {error && <p className="text-[11px] text-destructive">{error}</p>}
        {images && files === null && <p className="shimmer text-[11px]">Loading images…</p>}
        {files && files.length > 0 && (
          <div className="max-h-40 overflow-y-auto border-t pt-1">
            {files.map((file) => (
              <button
                key={file.id}
                type="button"
                disabled={busy}
                onClick={() => void run(() => images!.pick(file.id))}
                className="flex w-full items-center gap-2 rounded-md px-1.5 py-1 text-left text-[12px] hover:bg-muted"
              >
                <ImageIcon className="size-3.5 shrink-0 text-muted-foreground" />
                <span className="truncate">{file.name}</span>
              </button>
            ))}
          </div>
        )}
      </PopoverContent>
    </Popover>
  );
}
