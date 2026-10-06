import { useEffect, useState } from "react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { settings } from "@/lib/store";
import type { Settings } from "@/lib/store";

export function SettingsDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [value, setValue] = useState<Settings | null>(null);
  const [revealed, setRevealed] = useState(false);
  useEffect(() => {
    if (open) void settings.get().then(setValue);
    setRevealed(false);
  }, [open]);

  const save = async () => {
    if (value) await settings.set(value);
    onClose();
  };

  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="grid-cols-[minmax(0,1fr)] gap-5">
        <DialogTitle className="text-[13px] font-medium">Settings</DialogTitle>
        {value && (
          <>
            <label className="flex flex-col gap-1.5">
              <span className="text-[12px] font-medium">AI key</span>
              <Textarea
                rows={3}
                spellCheck={false}
                placeholder="sk-… or NAME=value lines"
                value={value.aiKey}
                onFocus={() => setRevealed(true)}
                onChange={(e) => setValue({ ...value, aiKey: e.target.value })}
                className="field-sizing-fixed resize-none font-mono text-[12px] break-all"
                style={{ WebkitTextSecurity: revealed ? "none" : "disc" } as React.CSSProperties}
              />
              <span className="text-[11px] text-muted-foreground">Stored in this browser only.</span>
            </label>
            <label className="flex items-center justify-between gap-4">
              <span className="flex flex-col">
                <span className="text-[12px] font-medium">Ask before changes</span>
                <span className="text-[11px] text-muted-foreground">
                  Otherwise changes are applied directly and can be undone.
                </span>
              </span>
              <Switch
                checked={value.approval === "ask"}
                onCheckedChange={(checked) => setValue({ ...value, approval: checked ? "ask" : "auto" })}
              />
            </label>
            <div className="flex justify-end gap-2">
              <Button variant="ghost" onClick={onClose}>
                Cancel
              </Button>
              <Button onClick={() => void save()}>Save</Button>
            </div>
          </>
        )}
      </DialogContent>
    </Dialog>
  );
}
