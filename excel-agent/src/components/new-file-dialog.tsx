import { useState } from "react";
import { useNavigate } from "@tanstack/react-router";
import { FilePlusIcon, FileSpreadsheetIcon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { handles, links, schemas } from "@/lib/store";
import { emptyWorkbook } from "@/storage/local";
import { createDriveFile } from "@/storage/sharepoint";
import type { DriveRef, FileSource } from "@/storage/types";
import { ImagesPicker } from "./images-picker";
import type { ImagesChoice } from "./images-picker";
import { SearchSelect } from "./search-select";
import { SimpleSelect } from "./simple-select";
import { DriveLinkField, MicrosoftSignIn, StorageTabs, useMicrosoft } from "./storage-ui";
import type { StorageKind } from "./storage-ui";

export const EXCEL: FilePickerAcceptType = {
  description: "Excel workbook",
  accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
};

/** Links an Excel file (local or SharePoint / OneDrive, existing or new) to a schema, with an optional images folder. */
export function NewFileDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const all = schemas.use();
  const navigate = useNavigate();
  const microsoft = useMicrosoft();
  const [schemaId, setSchemaId] = useState<string | null>(null);
  const [name, setName] = useState("");
  const [kind, setKind] = useState<StorageKind>("local");
  const [local, setLocal] = useState<FileSystemFileHandle>();
  const [remote, setRemote] = useState<{ mode: "existing" | "new"; file?: DriveRef; folder?: DriveRef; fileName: string }>({ mode: "existing", fileName: "" });
  const [images, setImages] = useState<ImagesChoice>();
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});

  const pick = async (create: boolean) => {
    try {
      const handle = create
        ? await window.showSaveFilePicker({ suggestedName: `${name || "data"}.xlsx`, types: [EXCEL] })
        : (await window.showOpenFilePicker({ types: [EXCEL] }))[0];
      setLocal(handle);
      if (!name) setName(handle.name.replace(/\.xlsx$/i, ""));
    } catch (e) {
      if ((e as Error).name !== "AbortError") setState({ error: (e as Error).message });
    }
  };

  const ready =
    !!schemaId && (kind === "local" ? !!local : remote.mode === "existing" ? !!remote.file : !!remote.folder && !!(remote.fileName || name).trim());

  const create = async () => {
    const schema = all.find((s) => s.id === schemaId);
    if (!schema || !ready) return;
    setState({ busy: true });
    try {
      const id = crypto.randomUUID();
      let source: FileSource;
      if (kind === "local") {
        await handles.setFile(id, local!);
        source = { kind: "local", name: local!.name };
      } else {
        const file = remote.mode === "existing" ? remote.file! : await createDriveFile(remote.folder!, (remote.fileName || name).trim(), await emptyWorkbook(schema.sheet));
        source = { kind: "sharepoint", ...file };
      }
      if (images?.kind === "local") await handles.setImages(id, images.handle);
      links.put({
        id,
        name: name.trim() || source.name.replace(/\.xlsx$/i, ""),
        schemaId: schema.id,
        source,
        images: images?.kind === "local" ? { kind: "local", name: images.handle.name } : images,
        openedAt: Date.now(),
      });
      close();
      void navigate({ to: "/files/$id", params: { id } });
    } catch (e) {
      setState({ error: e instanceof Error ? e.message : String(e) });
    }
  };

  const close = () => {
    setSchemaId(null);
    setName("");
    setLocal(undefined);
    setRemote({ mode: "existing", fileName: "" });
    setImages(undefined);
    setState({});
    onClose();
  };

  const localSupported = "showOpenFilePicker" in window;
  const signedIn = !!microsoft.status?.signedIn;

  return (
    <Dialog open={open} onOpenChange={(o) => !o && close()}>
      <DialogContent className="max-h-[90dvh] grid-cols-[minmax(0,1fr)] gap-4 overflow-y-auto sm:max-w-md">
        <DialogTitle className="text-[13px] font-medium">Add a file</DialogTitle>
        <Row label="Schema">
          {all.length === 0 ? (
            <p className="text-[12px] text-muted-foreground">Create a schema first (Schemas tab).</p>
          ) : (
            <SearchSelect
              value={schemaId}
              onChange={setSchemaId}
              onAdd={() => {
                close();
                void navigate({ to: "/schemas" });
              }}
              addLabel="New schema"
              options={all.map((s) => ({ value: s.id, label: s.name }))}
            />
          )}
        </Row>
        <Row label="Excel file">
          <StorageTabs value={kind} onChange={setKind} />
          {kind === "local" ? (
            !localSupported ? (
              <p className="text-[12px] text-destructive">This browser cannot open local files: use Chrome or Edge, or SharePoint.</p>
            ) : local ? (
              <div className="flex h-8 items-center gap-2 rounded-lg border bg-muted/40 px-2.5 [&_svg]:size-3.5 [&_svg]:text-muted-foreground">
                <FileSpreadsheetIcon />
                <span className="flex-1 truncate">{local.name}</span>
                <button type="button" aria-label="Clear" onClick={() => setLocal(undefined)}>
                  <XIcon />
                </button>
              </div>
            ) : (
              <div className="flex gap-2">
                <Button variant="outline" size="sm" className="flex-1" onClick={() => void pick(false)}>
                  <FileSpreadsheetIcon /> Choose a file…
                </Button>
                <Button variant="outline" size="sm" className="flex-1" onClick={() => void pick(true)}>
                  <FilePlusIcon /> New file…
                </Button>
              </div>
            )
          ) : (
            <>
              <MicrosoftSignIn onSignedIn={() => void microsoft.refresh()} />
              {signedIn && (
                <>
                  <SimpleSelect
                    size="sm"
                    value={remote.mode}
                    onChange={(mode) => setRemote({ ...remote, mode: mode as "existing" | "new" })}
                    options={[
                      { value: "existing", label: "Existing file (link)" },
                      { value: "new", label: "New file in a folder (link)" },
                    ]}
                  />
                  {remote.mode === "existing" ? (
                    <DriveLinkField
                      kind="file"
                      value={remote.file}
                      onChange={(file) => {
                        setRemote({ ...remote, file });
                        if (file && !name) setName(file.name.replace(/\.xlsx$/i, ""));
                      }}
                      placeholder="Link to the .xlsx (Share › Copy link)"
                    />
                  ) : (
                    <>
                      <DriveLinkField kind="folder" value={remote.folder} onChange={(folder) => setRemote({ ...remote, folder })} placeholder="Link to the folder" />
                      <Input value={remote.fileName} onChange={(e) => setRemote({ ...remote, fileName: e.target.value })} placeholder={`File name (default: ${name || "name below"})`} />
                    </>
                  )}
                </>
              )}
            </>
          )}
        </Row>
        <Row label="Name">
          <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Shown in the list" />
        </Row>
        <Row label="Images folder" hint="Optional: where image fields store their files (local names, or SharePoint links).">
          <ImagesPicker value={images} onChange={setImages} />
        </Row>
        {state.error && <p className="text-[12px] text-destructive">{state.error}</p>}
        <div className="flex justify-end gap-2">
          <Button variant="ghost" onClick={close}>
            Cancel
          </Button>
          <Button disabled={!ready || state.busy} onClick={() => void create()}>
            {state.busy ? "Creating…" : "Open"}
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

function Row({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1.5">
      <span className="text-[12px] font-medium">{label}</span>
      {children}
      {hint && <span className="text-[11px] text-muted-foreground">{hint}</span>}
    </div>
  );
}
