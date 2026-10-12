/** Editing one value by field type, in a table cell or in a form. */
import { useEffect, useRef, useState } from "react";
import { CheckIcon, ExternalLinkIcon, FolderIcon, ImageIcon, SearchIcon, UploadIcon } from "lucide-react";
import { imageUrl } from "@/lib/excel";
import { errorText } from "@/lib/ms";
import { TINTS, isIdField, listOf, type Field } from "@/lib/schema";
import { Markdown } from "./markdown";
import { Popover } from "./popover";
import { RichText } from "./rich-text";

/** Image uploads: only into a SharePoint folder the user chose (`folder` = its name). */
export type Uploader = { save: (file: File) => Promise<string>; folder?: string; chooseFolder: (link: string) => Promise<unknown> };

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
  // Ids are given by the app: read only.
  if (isIdField(props.field))
    return <span className={`flex h-8 items-center text-[12px] text-muted-foreground tabular-nums ${props.variant === "cell" ? "px-2" : ""}`}>{String(props.value ?? "")}</span>;
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
function TextInput({ field, value, onChange, variant, placeholder }: InputProps & { placeholder?: string }) {
  const shown = value === null || value === undefined ? "" : String(value);
  const [draft, setDraft] = useState(shown);
  useEffect(() => setDraft(shown), [shown]);
  const numeric = field.type === "number" || field.type === "integer";
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

/** Rich text stored as markdown: full editor in a form, preview + editor in a popover in a cell. */
function MarkdownInput({ value, onChange, variant }: InputProps) {
  const shown = String(value ?? "");
  if (variant === "form") return <RichText value={shown} onChange={onChange} />;
  return (
    <Popover full width={480} trigger={(open) => (
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
    <Popover full width={220} trigger={(open) => (
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

/** Image field = a link column: paste a URL, or upload a file to the chosen SharePoint folder (its link is stored). */
function ImageInput(props: InputProps) {
  const { value, variant } = props;
  const url = String(value ?? "");
  return (
    <div className={`flex items-center ${variant === "form" ? "gap-1" : "pr-1"}`}>
      <span className="min-w-0 flex-1">
        <TextInput {...props} placeholder={variant === "form" ? "https://… (image link)" : undefined} />
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
  );
}

function UploadButton({ onChange, upload }: InputProps) {
  const file = useRef<HTMLInputElement>(null);
  const [folder, setFolder] = useState<string>();
  const [state, setState] = useState<{ busy?: string; error?: string }>({});
  if (!upload) return null;
  const run = (busy: string, task: Promise<unknown>, done?: () => void) => {
    setState({ busy });
    task.then(() => (setState({}), done?.()), (e) => setState({ error: errorText(e) }));
  };
  const choosing = !upload.folder || folder !== undefined;
  return (
    <Popover width={300} trigger={(open) => (
      <button type="button" className="icon-btn" title="Upload an image" onClick={() => (setFolder(undefined), setState({}), open())}>
        <UploadIcon />
      </button>
    )}>
      {(close) =>
        choosing ? (
          <div className="flex flex-col gap-2 p-1">
            <span className="text-[12px]">{upload.folder ? "Other folder for the uploads" : "First choose the SharePoint folder that receives the images"}: paste its link.</span>
            <input autoFocus className="input h-7 text-[12px]" placeholder="https://…sharepoint.com/:f:/…" value={folder ?? ""} onChange={(e) => setFolder(e.target.value)} />
            <div className="flex justify-end gap-1.5">
              {upload.folder && (
                <button type="button" className="btn h-7" onClick={() => setFolder(undefined)}>
                  Back
                </button>
              )}
              <button type="button" className="btn-primary h-7" disabled={!folder?.trim() || !!state.busy} onClick={() => run("Checking…", upload.chooseFolder(folder!.trim()), () => setFolder(undefined))}>
                <FolderIcon /> {state.busy ?? "Use this folder"}
              </button>
            </div>
            {state.error && <p className="text-[11px] text-destructive">{state.error}</p>}
          </div>
        ) : (
          <div className="flex flex-col gap-2 p-1">
            <input
              ref={file}
              type="file"
              accept="image/*"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0];
                e.target.value = "";
                if (f) run("Uploading…", upload.save(f).then(onChange), close);
              }}
            />
            <button type="button" className="btn-primary h-7" disabled={!!state.busy} onClick={() => file.current?.click()}>
              <UploadIcon /> {state.busy ?? "Choose an image"}
            </button>
            <p className="text-[11px] text-muted-foreground">
              Goes to the folder "{upload.folder}"; its link is stored in the cell.{" "}
              <button type="button" className="underline" onClick={() => setFolder("")}>
                Change the folder
              </button>
            </p>
            {state.error && <p className="text-[11px] text-destructive">{state.error}</p>}
          </div>
        )
      }
    </Popover>
  );
}
