/** Editing one value by field type, in a table cell or in a form. */
import { useEffect, useRef, useState } from "react";
import { CheckIcon, ImageIcon, SearchIcon, UploadIcon, XIcon } from "lucide-react";
import { imageUrl } from "@/lib/excel";
import { errorText } from "@/lib/ms";
import { TINTS, listOf, type Field } from "@/lib/schema";
import { Markdown } from "./markdown";
import { Popover } from "./popover";
import { RichText } from "./rich-text";

export type Uploader = { save: (file: File) => Promise<string>; label: string };

export type InputProps = {
  field: Field;
  value: unknown;
  onChange: (value: unknown) => void;
  variant: "cell" | "form";
  /** Where image fields upload their files. */
  upload?: Uploader;
};

const base = {
  cell: "h-8 w-full rounded-sm bg-transparent px-2 outline-none focus:bg-card focus:ring-1 focus:ring-ring",
  form: "input",
};

export function FieldInput(props: InputProps) {
  switch (props.field.type) {
    case "boolean":
      return (
        <div className={`flex h-8 items-center ${props.variant === "cell" ? "px-2" : ""}`}>
          <input type="checkbox" className="size-4 accent-foreground" checked={props.value === true} onChange={(e) => props.onChange(e.target.checked)} />
        </div>
      );
    case "option":
      return <OptionInput {...props} />;
    case "image":
      return <ImageInput {...props} />;
    case "text":
      return <MarkdownInput {...props} />;
    default:
      return <TextInput {...props} />;
  }
}

/** Committed on blur / Enter, so each keystroke does not write the file. */
function TextInput({ field, value, onChange, variant }: InputProps) {
  const shown = value === null || value === undefined ? "" : String(value);
  const [draft, setDraft] = useState(shown);
  useEffect(() => setDraft(shown), [shown]);
  const numeric = field.type === "number" || field.type === "integer";
  return (
    <input
      type={field.type === "date" ? "date" : "text"}
      inputMode={numeric ? "decimal" : undefined}
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => draft !== shown && onChange(draft)}
      onKeyDown={(e) => {
        if (e.key === "Enter") e.currentTarget.blur();
        if (e.key === "Escape") setDraft(shown);
      }}
      className={`${base[variant]} ${numeric ? "text-right tabular-nums" : ""}`}
    />
  );
}

/** Rich text stored as markdown: full editor in a form, preview + editor in a popover in a cell. */
function MarkdownInput({ value, onChange, variant }: InputProps) {
  const shown = String(value ?? "");
  if (variant === "form") return <RichText value={shown} onChange={onChange} />;
  return (
    <Popover width={480} trigger={(open) => (
      <button type="button" onClick={open} className="flex min-h-8 w-full items-start px-2 py-1.5 text-left">
        <span className="line-clamp-2 text-[12px]">{shown ? <Markdown>{shown}</Markdown> : ""}</span>
      </button>
    )}>
      {(close) => (
        <>
          <RichText value={shown} onChange={onChange} autoFocus />
          <button type="button" className="btn-primary h-7 self-end" onClick={close}>
            Done
          </button>
        </>
      )}
    </Popover>
  );
}

export function OptionBadges({ field, value }: { field: Field; value: unknown }) {
  return (
    <>
      {listOf(value).map((v) => {
        const option = field.options?.find((o) => o.value === v);
        return (
          <span key={v} className={`inline-flex h-5 items-center rounded-full px-2 text-[12px] whitespace-nowrap ${option ? TINTS[option.color] : "bg-red-500/10 text-destructive"}`}>
            {v}
          </span>
        );
      })}
    </>
  );
}

function OptionInput({ field, value, onChange, variant }: InputProps) {
  const [search, setSearch] = useState("");
  const selected = listOf(value);
  const options = (field.options ?? []).filter((o) => o.value.toLowerCase().includes(search.toLowerCase()));
  const toggle = (v: string, close: () => void) => {
    if (field.multiple) return onChange(selected.includes(v) ? selected.filter((s) => s !== v) : [...selected, v]);
    onChange(selected[0] === v ? null : v);
    close();
  };
  return (
    <Popover width={220} trigger={(open) => (
      <button type="button" onClick={open} className={`flex min-h-8 w-full flex-wrap items-center gap-1 text-left ${variant === "cell" ? "px-2 py-1" : "rounded-md border bg-card px-2 py-1"}`}>
        <OptionBadges field={field} value={value} />
        {selected.length === 0 && variant === "form" && <span className="text-muted-foreground">Select…</span>}
      </button>
    )}>
      {(close) => (
        <>
          {(field.options?.length ?? 0) > 6 && (
            <label className="flex items-center gap-1.5 border-b px-1 pb-1">
              <SearchIcon className="size-3.5 text-muted-foreground" />
              <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…" className="h-7 flex-1 bg-transparent outline-none" />
            </label>
          )}
          <div className="max-h-64 overflow-y-auto">
            {options.map((o) => (
              <button key={o.value} type="button" onClick={() => toggle(o.value, close)} className="flex w-full items-center gap-2 rounded-md px-2 py-1 text-left hover:bg-muted">
                <span className={`inline-flex h-5 items-center rounded-full px-2 text-[12px] ${TINTS[o.color]}`}>{o.value}</span>
                {selected.includes(o.value) && <CheckIcon className="ml-auto size-3.5" />}
              </button>
            ))}
            {options.length === 0 && <p className="px-2 py-1 text-[12px] text-muted-foreground">No option: add some in Fields.</p>}
          </div>
        </>
      )}
    </Popover>
  );
}

export function Thumbnail({ value, className = "" }: { value: unknown; className?: string }) {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => void imageUrl(value).then(setUrl), [value]);
  if (!value) return null;
  return url ? (
    <img src={url} alt="" className={`rounded object-cover ring-1 ring-black/10 ${className}`} />
  ) : (
    <span title={String(value)} className={`inline-flex items-center justify-center rounded bg-muted text-muted-foreground ${className}`}>
      <ImageIcon className="size-3.5" />
    </span>
  );
}

function ImageInput({ value, onChange, variant, upload }: InputProps) {
  const file = useRef<HTMLInputElement>(null);
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const send = (f: File, close: () => void) => {
    setState({ busy: true });
    upload!.save(f).then(
      (url) => {
        onChange(url);
        setState({});
        close();
      },
      (e) => setState({ error: errorText(e) }),
    );
  };
  const picker = (
    <Popover width={320} trigger={(open) => (
      <button type="button" onClick={open} className={`flex min-h-8 w-full items-center gap-2 text-left ${variant === "cell" ? "px-2 py-1" : "rounded-md border bg-card px-2 py-1.5"}`}>
        {variant === "cell" ? <Thumbnail value={value} className="size-7" /> : <ImageIcon className="size-4 shrink-0 text-muted-foreground" />}
        {variant === "form" && <span className="truncate text-[12px] text-muted-foreground">{value ? String(value) : "No image: click to upload or paste a link"}</span>}
      </button>
    )}>
      {(close) => (
        <>
          <div className="flex gap-1.5">
            {upload && (
              <>
                <input ref={file} type="file" accept="image/*" hidden onChange={(e) => e.target.files?.[0] && send(e.target.files[0], close)} />
                <button type="button" className="btn h-7 flex-1" disabled={state.busy} onClick={() => file.current?.click()}>
                  <UploadIcon /> {state.busy ? "Uploading…" : "Upload"}
                </button>
              </>
            )}
            {!!value && (
              <button type="button" className="btn h-7" onClick={() => (onChange(null), close())}>
                <XIcon /> Clear
              </button>
            )}
          </div>
          {upload && <p className="text-[11px] text-muted-foreground">Uploaded to {upload.label}</p>}
          <input
            value={link}
            onChange={(e) => setLink(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && /^https?:\/\//i.test(link.trim()) && (onChange(link.trim()), setLink(""), close())}
            placeholder="Or paste an image link (Enter)"
            className="input h-7 text-[12px]"
          />
          {state.error && <p className="text-[11px] text-destructive">{state.error}</p>}
        </>
      )}
    </Popover>
  );
  if (variant === "cell") return picker;
  return (
    <div className="flex flex-col gap-2">
      {picker}
      {!!value && <Thumbnail value={value} className="max-h-96 w-full object-contain bg-muted" />}
    </div>
  );
}
