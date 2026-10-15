/** Editing one value by field type, in a table cell or in a form; and the badges of choices. */
import { useEffect, useRef, useState } from "react";
import { CheckIcon, ExternalLinkIcon, FileIcon, SearchIcon, StarIcon, UploadIcon } from "lucide-react";
import { imageUrl } from "@/lib/excel";
import { errorText } from "@/lib/ms";
import { SWATCHES, TINTS, isIdField, isImageUrl, listOf, type Color, type Field } from "@/lib/schema";
import { Markdown } from "./markdown";
import { Popover } from "./popover";
import { RichText } from "./rich-text";

/** File uploads: into the schema's uploads folder (`folder` = its name), set in the Schema editor. */
export type Uploader = { save: (file: File) => Promise<string>; folder: string };

export type InputProps = {
  field: Field;
  value: unknown;
  onChange: (value: unknown) => void;
  variant: "cell" | "form";
  /** Where image and file fields upload their files (none: links only). */
  upload?: Uploader;
  /** Id fields: why a new value is refused (e.g. already used); undefined = accepted. */
  validate?: (value: string) => string | undefined;
  readOnly?: boolean;
};

const base = {
  cell: "h-11 w-full rounded-md bg-transparent px-4 outline-none focus:bg-card focus:ring-2 focus:ring-zinc-900/10",
  form: "input",
};

export function FieldInput(props: InputProps) {
  if (props.readOnly) return <FieldValue field={props.field} value={props.value} className={props.variant === "cell" ? "px-4 py-2.5" : ""} />;
  // Ids are given by the app; the user may rename one (checked by `validate`).
  if (isIdField(props.field)) return <IdInput {...props} />;
  switch (props.field.type) {
    case "option":
      return <OptionInput {...props} />;
    case "rating":
      return <RatingInput {...props} />;
    case "image":
    case "file":
      return <FileInput {...props} />;
    case "url":
      return <UrlInput {...props} />;
    case "text":
      return <MarkdownInput {...props} />;
    default:
      return <TextInput {...props} />;
  }
}

/** A value shown (not edited), as in the table: badges, stars, thumbnails, markdown. */
export function FieldValue({ field, value, className = "" }: { field: Field; value: unknown; className?: string }) {
  if (value === null || value === undefined || value === "" || (Array.isArray(value) && !value.length)) return <span className={`text-muted-foreground/60 ${className}`}>—</span>;
  switch (field.type) {
    case "option":
      return (
        <span className={`flex flex-wrap gap-1.5 ${className}`}>
          <OptionBadges field={field} value={value} />
        </span>
      );
    case "rating":
      return <Stars max={field.max ?? 5} value={Number(value)} className={className} />;
    case "image":
    case "file":
      return field.type === "image" || isImageUrl(value) ? <Thumbnail value={value} className={`size-9 ${className}`} /> : <FileLink value={value} className={className} />;
    case "url":
      return (
        <a href={String(value)} target="_blank" rel="noreferrer" className={`truncate text-[13.5px] hover:underline ${className}`}>
          {String(value).replace(/^https?:\/\//, "")}
        </a>
      );
    case "text":
      return (
        <span className={`line-clamp-2 ${className}`}>
          <Markdown>{String(value)}</Markdown>
        </span>
      );
    default:
      return <span className={`${isIdField(field) ? "text-muted-foreground tabular-nums" : ""} ${className}`}>{String(value)}</span>;
  }
}

/** Committed on blur / Enter, so each keystroke does not write the file. */
function TextInput({ field, value, onChange, variant, placeholder }: InputProps & { placeholder?: string }) {
  const shown = value === null || value === undefined ? "" : String(value);
  const [draft, setDraft] = useState(shown);
  useEffect(() => setDraft(shown), [shown]);
  const numeric = field.type === "number";
  return (
    <input
      type={field.type === "date" ? "date" : "text"}
      inputMode={numeric ? "decimal" : undefined}
      value={draft}
      placeholder={placeholder}
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

/** Row id: committed on blur / Enter only when `validate` accepts it, else the reason is shown (Escape restores). */
function IdInput({ value, onChange, variant, validate }: InputProps) {
  const shown = String(value ?? "");
  const [draft, setDraft] = useState(shown);
  const [error, setError] = useState<string>();
  useEffect(() => (setDraft(shown), setError(undefined)), [shown]);
  const commit = () => {
    const next = draft.trim();
    if (next === shown) return setError(undefined);
    const reason = validate?.(next);
    if (reason) return setError(reason);
    setError(undefined);
    onChange(next);
  };
  return (
    <div className="flex flex-col gap-1">
      <input
        value={draft}
        title={error}
        aria-invalid={!!error}
        onChange={(e) => (setDraft(e.target.value), setError(undefined))}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") e.currentTarget.blur();
          if (e.key === "Escape") (setDraft(shown), setError(undefined));
        }}
        className={`${base[variant]} text-muted-foreground tabular-nums focus:text-foreground ${error ? "text-destructive ring-1 ring-destructive" : ""}`}
      />
      {error && <span className={`text-[12px] text-destructive ${variant === "cell" ? "px-4 pb-1" : ""}`}>{error}</span>}
    </div>
  );
}

/** Long text stored as markdown: full editor in a form, preview + editor in a popover in a cell. */
function MarkdownInput({ value, onChange, variant }: InputProps) {
  const shown = String(value ?? "");
  if (variant === "form") return <RichText value={shown} onChange={onChange} />;
  return (
    <Popover full width={520} trigger={(open) => (
      <button type="button" onClick={open} className="flex min-h-11 w-full items-start px-4 py-2.5 text-left">
        <span className="line-clamp-2 text-[13.5px]">{shown ? <Markdown>{shown}</Markdown> : ""}</span>
      </button>
    )}>
      {(close) => (
        <>
          <RichText value={shown} onChange={onChange} autoFocus />
          <button type="button" className="btn-primary h-8 self-end" onClick={close}>
            Done
          </button>
        </>
      )}
    </Popover>
  );
}

/** A choice as a badge: tinted background, colored dot. */
export const Badge = ({ color, small, children }: { color?: Color; small?: boolean; children: string }) => (
  <span title={children} className={`badge max-w-full min-w-0 ${small ? "h-5.5 gap-1 px-2 text-[11.5px]" : ""} ${color ? TINTS[color] : "bg-red-500/10 text-destructive"}`}>
    <span className={`size-1.5 shrink-0 rounded-full ${color ? SWATCHES[color] : "bg-destructive"}`} />
    <span className="truncate">{children}</span>
  </span>
);

export function OptionBadges({ field, value }: { field: Field; value: unknown }) {
  return (
    <>
      {listOf(value).map((v) => (
        <Badge key={v} color={field.options?.find((o) => o.value === v)?.color}>
          {v}
        </Badge>
      ))}
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
    <Popover full width={240} trigger={(open) => (
      <button type="button" onClick={open} className={`flex min-h-11 w-full min-w-0 flex-wrap items-center gap-1.5 overflow-hidden text-left ${variant === "cell" ? "px-4 py-2" : "rounded-lg border bg-card px-2.5 py-1.5"}`}>
        <OptionBadges field={field} value={value} />
        {selected.length === 0 && variant === "form" && <span className="text-muted-foreground">Select…</span>}
      </button>
    )}>
      {(close) => (
        <>
          {(field.options?.length ?? 0) > 6 && (
            <label className="flex items-center gap-2 border-b px-1.5 pb-1.5">
              <SearchIcon className="size-4 text-muted-foreground" />
              <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…" className="h-8 flex-1 bg-transparent outline-none" />
            </label>
          )}
          <div className="max-h-72 overflow-y-auto">
            {options.map((o) => (
              <button key={o.value} type="button" onClick={() => toggle(o.value, close)} className="flex w-full items-center gap-2 rounded-lg px-2 py-1.5 text-left hover:bg-muted">
                <Badge color={o.color}>{o.value}</Badge>
                {selected.includes(o.value) && <CheckIcon className="ml-auto size-4" />}
              </button>
            ))}
            {options.length === 0 && <p className="px-2 py-1.5 text-[13px] text-muted-foreground">No choice: add some in the Schema editor.</p>}
          </div>
        </>
      )}
    </Popover>
  );
}

const Stars = ({ max, value, onPick, className = "" }: { max: number; value: number; onPick?: (n: number | null) => void; className?: string }) => (
  <span className={`flex items-center gap-0.5 ${className}`}>
    {Array.from({ length: max }, (_, i) => (
      <button key={i} type="button" disabled={!onPick} onClick={() => onPick?.(value === i + 1 ? null : i + 1)} className="disabled:cursor-default" title={`${i + 1}`}>
        <StarIcon className={`size-4 ${i < value ? "fill-amber-400 text-amber-400" : "text-zinc-300"}`} />
      </button>
    ))}
  </span>
);

function RatingInput({ field, value, onChange, variant }: InputProps) {
  return <Stars max={field.max ?? 5} value={Number(value) || 0} onPick={onChange} className={variant === "cell" ? "h-11 px-4" : "h-9"} />;
}

export function Thumbnail({ value, className = "" }: { value: unknown; className?: string }) {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => void imageUrl(value).then(setUrl), [value]);
  if (!value) return null;
  return url ? (
    <img src={url} alt="" className={`rounded-md object-cover ring-1 ring-black/10 ${className}`} />
  ) : (
    <span title={String(value)} className={`inline-flex items-center justify-center rounded-md bg-muted text-muted-foreground ${className}`}>
      <FileIcon className="size-4" />
    </span>
  );
}

const FileLink = ({ value, className = "" }: { value: unknown; className?: string }) => (
  <a href={String(value)} target="_blank" rel="noreferrer" className={`flex min-w-0 items-center gap-1.5 text-[13.5px] hover:underline ${className}`}>
    <FileIcon className="size-4 shrink-0 text-muted-foreground" />
    <span className="truncate">{decodeURIComponent(String(value).split("/").pop()?.split("?")[0] ?? "file")}</span>
  </a>
);

/** Link field: the URL, and a button opening it. */
function UrlInput(props: InputProps) {
  const url = String(props.value ?? "");
  return (
    <div className={`flex items-center gap-1 ${props.variant === "cell" ? "pr-1" : ""}`}>
      <span className="min-w-0 flex-1">
        <TextInput {...props} placeholder={props.variant === "form" ? "https://…" : undefined} />
      </span>
      {/^https?:\/\//i.test(url) && (
        <a className={`icon-btn shrink-0 ${props.variant === "cell" ? "opacity-0 group-hover:opacity-100" : ""}`} href={url} target="_blank" rel="noreferrer" title="Open the link">
          <ExternalLinkIcon />
        </a>
      )}
    </div>
  );
}

/**
 * Image and file fields = a link column: paste a URL, or upload a file (when the schema has an uploads
 * folder: its link is stored). An image shows in the form.
 */
function FileInput(props: InputProps) {
  const { field, value, variant } = props;
  const url = String(value ?? "");
  const image = field.type === "image" || isImageUrl(url);
  return (
    <div className="flex flex-col gap-2">
      {variant === "form" && field.type === "image" && url && <Thumbnail value={url} className="max-h-64 w-fit max-w-full object-contain!" />}
      <div className={`flex items-center gap-1 ${variant === "cell" ? "pr-1" : ""}`}>
        {variant === "cell" && image && url && <Thumbnail value={url} className="ml-3 size-8 shrink-0" />}
        <span className="min-w-0 flex-1">
          <TextInput {...props} placeholder={variant === "form" ? `https://… (${field.type === "image" ? "image" : "file"} link)` : undefined} />
        </span>
        <span className={`flex shrink-0 ${variant === "cell" ? "opacity-0 group-hover:opacity-100 focus-within:opacity-100" : ""}`}>
          {/^https?:\/\//i.test(url) && (
            <a className="icon-btn" href={url} target="_blank" rel="noreferrer" title="Open the link">
              <ExternalLinkIcon />
            </a>
          )}
          <UploadButton {...props} />
        </span>
      </div>
    </div>
  );
}

/** Uploads one file to the schema's uploads folder and stores its link. */
function UploadButton({ field, onChange, upload }: InputProps) {
  const file = useRef<HTMLInputElement>(null);
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  if (!upload) return null;
  return (
    <>
      <input
        ref={file}
        type="file"
        hidden
        accept={field.type === "image" ? "image/*" : undefined}
        onChange={(e) => {
          const f = e.target.files?.[0];
          e.target.value = "";
          if (!f) return;
          setState({ busy: true });
          upload.save(f).then((url) => (setState({}), onChange(url)), (err) => setState({ error: errorText(err) }));
        }}
      />
      <button type="button" className="icon-btn" disabled={state.busy} title={state.error ?? `Upload a ${field.type === "image" ? "image" : "file"} (to "${upload.folder}")`} onClick={() => file.current?.click()}>
        <UploadIcon className={state.busy ? "animate-pulse" : state.error ? "text-destructive" : ""} />
      </button>
    </>
  );
}
