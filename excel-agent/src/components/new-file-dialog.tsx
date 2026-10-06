import { useState } from "react";
import { useNavigate } from "@tanstack/react-router";
import { FilePlusIcon, FileSpreadsheetIcon, FolderIcon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { handles, links, schemas } from "@/lib/store";
import { SimpleSelect } from "./simple-select";

const EXCEL: FilePickerAcceptType = {
  description: "Excel workbook",
  accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
};

/** Links a local Excel file (existing or new) to a schema, with an optional images folder. */
export function NewFileDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const all = schemas.use();
  const navigate = useNavigate();
  const [schemaId, setSchemaId] = useState<string | null>(null);
  const [name, setName] = useState("");
  const [file, setFile] = useState<FileSystemFileHandle>();
  const [images, setImages] = useState<FileSystemDirectoryHandle>();
  const [error, setError] = useState("");

  const pick = async (create: boolean) => {
    setError("");
    try {
      const handle = create
        ? await window.showSaveFilePicker({ suggestedName: `${name || "data"}.xlsx`, types: [EXCEL] })
        : (await window.showOpenFilePicker({ types: [EXCEL] }))[0];
      setFile(handle);
      if (!name) setName(handle.name.replace(/\.xlsx$/i, ""));
    } catch (e) {
      if ((e as Error).name !== "AbortError") setError((e as Error).message);
    }
  };

  const pickFolder = async () => {
    try {
      setImages(await window.showDirectoryPicker({ mode: "readwrite" }));
    } catch (e) {
      if ((e as Error).name !== "AbortError") setError((e as Error).message);
    }
  };

  const create = async () => {
    if (!file || !schemaId) return;
    const id = crypto.randomUUID();
    await handles.setFile(id, file);
    await handles.setImages(id, images);
    links.put({ id, name: name.trim() || file.name, schemaId, fileName: file.name, imagesFolder: images?.name, openedAt: Date.now() });
    close();
    void navigate({ to: "/files/$id", params: { id } });
  };

  const close = () => {
    setSchemaId(null);
    setName("");
    setFile(undefined);
    setImages(undefined);
    setError("");
    onClose();
  };

  const supported = "showOpenFilePicker" in window;

  return (
    <Dialog open={open} onOpenChange={(o) => !o && close()}>
      <DialogContent className="gap-4 sm:max-w-md">
        <DialogTitle className="text-[13px] font-medium">Add a file</DialogTitle>
        {!supported && (
          <p className="text-[12px] text-destructive">This browser cannot open local files: use Chrome or Edge.</p>
        )}
        <Row label="Schema">
          {all.length === 0 ? (
            <p className="text-[12px] text-muted-foreground">Create a schema first (Schemas tab).</p>
          ) : (
            <SimpleSelect value={schemaId} onChange={setSchemaId} options={all.map((s) => ({ value: s.id, label: s.name }))} />
          )}
        </Row>
        <Row label="Excel file">
          {file ? (
            <Chosen icon={<FileSpreadsheetIcon />} name={file.name} onClear={() => setFile(undefined)} />
          ) : (
            <div className="flex gap-2">
              <Button variant="outline" size="sm" className="flex-1" disabled={!supported} onClick={() => void pick(false)}>
                <FileSpreadsheetIcon /> Choose a file…
              </Button>
              <Button variant="outline" size="sm" className="flex-1" disabled={!supported} onClick={() => void pick(true)}>
                <FilePlusIcon /> New file…
              </Button>
            </div>
          )}
        </Row>
        <Row label="Name">
          <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Shown in the list" />
        </Row>
        <Row label="Images folder" hint="Optional: where image fields store their files.">
          {images ? (
            <Chosen icon={<FolderIcon />} name={images.name} onClear={() => setImages(undefined)} />
          ) : (
            <Button variant="outline" size="sm" className="w-fit" disabled={!supported} onClick={() => void pickFolder()}>
              <FolderIcon /> Choose a folder…
            </Button>
          )}
        </Row>
        {error && <p className="text-[12px] text-destructive">{error}</p>}
        <div className="flex justify-end gap-2">
          <Button variant="ghost" onClick={close}>
            Cancel
          </Button>
          <Button disabled={!file || !schemaId} onClick={() => void create()}>
            Open
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

function Chosen({ icon, name, onClear }: { icon: React.ReactNode; name: string; onClear: () => void }) {
  return (
    <div className="flex h-8 items-center gap-2 rounded-lg border bg-muted/40 px-2.5 [&_svg]:size-3.5 [&_svg]:text-muted-foreground">
      {icon}
      <span className="flex-1 truncate">{name}</span>
      <button type="button" aria-label="Clear" onClick={onClear} className="hover:text-foreground">
        <XIcon />
      </button>
    </div>
  );
}
