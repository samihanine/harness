import { useEffect, useRef, useState } from "react";
import { CheckIcon, ImageIcon, SearchIcon, UploadIcon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { TINTS } from "@/lib/colors";
import { imageUrl, listImages, saveImage } from "@/lib/images";
import type { Field } from "@/lib/schema";
import { cn } from "@/lib/utils";

export type InputProps = {
  field: Field;
  value: unknown;
  onChange: (value: unknown) => void;
  /** "cell": borderless, inside the table. "form": regular input. */
  variant: "cell" | "form";
  images?: FileSystemDirectoryHandle;
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

function TextArea({ value, onChange }: InputProps) {
  const shown = String(value ?? "");
  const [draft, setDraft] = useState(shown);
  useEffect(() => setDraft(shown), [shown]);
  return (
    <textarea
      value={draft}
      rows={4}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => draft !== shown && onChange(draft)}
      className="field-sizing-content min-h-20 w-full rounded-lg border bg-transparent px-2.5 py-1.5 outline-none transition-shadow focus:border-ring focus:ring-3 focus:ring-ring/30"
    />
  );
}

/** Long text in a cell: two lines, edited in a popover. */
function TextCell(props: InputProps) {
  return (
    <Popover>
      <PopoverTrigger className="flex min-h-8 w-full items-start px-2 py-1.5 text-left">
        <span className="line-clamp-2 whitespace-pre-wrap">{String(props.value ?? "")}</span>
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

function Thumbnail({ folder, name, className }: { folder?: FileSystemDirectoryHandle; name: unknown; className?: string }) {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => {
    void imageUrl(folder, name).then(setUrl);
  }, [folder, name]);
  if (!name) return null;
  return url ? (
    <img src={url} alt="" className={cn("rounded object-cover ring-1 ring-foreground/10", className)} />
  ) : (
    <span className={cn("inline-flex items-center justify-center rounded bg-muted text-muted-foreground", className)} title={String(name)}>
      <ImageIcon className="size-3.5" />
    </span>
  );
}

function ImageInput({ value, onChange, variant, images }: InputProps) {
  const input = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<string[] | null>(null);
  const upload = async (file: File) => images && onChange(await saveImage(images, file));
  return (
    <Popover onOpenChange={(open) => open && images && void listImages(images).then(setFiles)}>
      <PopoverTrigger className={cn("flex min-h-8 w-full items-center gap-2 text-left", variant === "cell" ? "px-2" : "rounded-lg border px-2 py-1")}>
        <Thumbnail folder={images} name={value} className={variant === "cell" ? "size-6" : "size-12"} />
        {variant === "form" && <span className="truncate text-muted-foreground">{value ? String(value) : "No image"}</span>}
      </PopoverTrigger>
      <PopoverContent align="start" className="w-64 gap-2 p-2">
        {value ? <Thumbnail folder={images} name={value} className="max-h-48 w-full" /> : null}
        {!images ? (
          <p className="text-[12px] text-muted-foreground">No images folder for this file: choose one with the folder button in the toolbar.</p>
        ) : (
          <>
            <input ref={input} type="file" accept="image/*" hidden onChange={(e) => e.target.files?.[0] && void upload(e.target.files[0])} />
            <div className="flex gap-1.5">
              <Button size="sm" variant="outline" className="flex-1" onClick={() => input.current?.click()}>
                <UploadIcon /> Upload
              </Button>
              {!!value && (
                <Button size="sm" variant="ghost" onClick={() => onChange(null)}>
                  <XIcon /> Clear
                </Button>
              )}
            </div>
            {files && files.length > 0 && (
              <div className="max-h-40 overflow-y-auto border-t pt-1">
                {files.map((name) => (
                  <button key={name} type="button" onClick={() => onChange(name)} className="flex w-full items-center gap-2 rounded-md px-1.5 py-1 text-left text-[12px] hover:bg-muted">
                    <Thumbnail folder={images} name={name} className="size-5" />
                    <span className="truncate">{name}</span>
                  </button>
                ))}
              </div>
            )}
          </>
        )}
      </PopoverContent>
    </Popover>
  );
}
