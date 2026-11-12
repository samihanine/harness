/**
 * Settings › Slides: the template deck — an ordinary PowerPoint where each slide is a type of slide
 * the assistant may add to any deck (name it in the slide's notes: "type: Two columns"). Kept in this
 * browser.
 */
import { useEffect, useRef, useState } from "react";
import { DownloadIcon, Trash2Icon, UploadIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ScrollPage, Section, StatusText } from "@/components/workspace";
import { download, keepTemplate, readTemplate, type TemplateFile } from "./files";
import { useRendered } from "./rendered";
import { Deck, PPTX_TYPE } from "./pptx/deck";
import { slideTypes, type SlideType } from "./pptx/template";

export default function TemplateSettings() {
  const [file, setFile] = useState<TemplateFile>();
  const [deck, setDeck] = useState<Deck>();
  const [types, setTypes] = useState<SlideType[]>([]);
  const [state, setState] = useState<{ busy?: string; done?: string; error?: string }>({});
  const input = useRef<HTMLInputElement>(null);
  const rendered = useRendered(deck, file?.at ?? 0, 1, "template");
  const show = async (t?: TemplateFile) => {
    setFile(t);
    const d = t ? await Deck.open(t.bytes) : undefined;
    setDeck(d);
    setTypes(d ? await slideTypes(d) : []);
  };
  useEffect(() => void readTemplate().then(show), []);
  const upload = async (f: File) => {
    setState({ busy: "Reading the deck…" });
    try {
      const t = { name: f.name, bytes: new Uint8Array(await f.arrayBuffer()), at: Date.now() };
      await show(t);
      await keepTemplate(t);
      setState({ done: "Template saved in this browser." });
    } catch (e) {
      setState({ error: e instanceof Error ? e.message : String(e) });
    }
  };
  return (
    <ScrollPage>
      <Section
        title="Template deck"
        description="A PowerPoint where each slide is a type of slide (cover, agenda, two columns, a chart page…). The assistant copies them, with their look, into the decks it edits and fills their texts. Name a type in the slide's notes (“type: Two columns”); otherwise its title names it."
        actions={
          <>
            <input ref={input} type="file" accept=".pptx" hidden onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ""; if (f) void upload(f); }} />
            <Button size="sm" variant="outline" onClick={() => input.current?.click()}>
              <UploadIcon /> {file ? "Replace" : "Upload a .pptx"}
            </Button>
            {file && deck && (
              <>
                <Button size="sm" variant="outline" onClick={() => download(new Blob([file.bytes as BlobPart], { type: PPTX_TYPE }), file.name)}>
                  <DownloadIcon /> Download
                </Button>
                <Button size="sm" variant="ghost" onClick={() => void keepTemplate(undefined).then(() => show(undefined))}>
                  <Trash2Icon /> Remove
                </Button>
              </>
            )}
          </>
        }
      >
        <StatusText state={state} />
        {file && <p className="text-[13px] text-muted-foreground">{file.name} · {types.length} slide types</p>}
        {deck && (
          <div className="mt-2 grid gap-4 sm:grid-cols-2">
            {types.map((t) => {
              return (
                <div key={t.number} className="flex flex-col gap-1.5">
                  {rendered.small[t.number] ? <img src={rendered.small[t.number]} alt={t.name} className="block w-full rounded-md shadow-sm ring-1 ring-border" /> : <div style={{ aspectRatio: `${deck.size.w} / ${deck.size.h}` }} className="animate-pulse rounded-md bg-muted/60 ring-1 ring-border" />}
                  <span className="text-[13px] font-medium">{t.name}</span>
                  <span className="text-[11.5px] text-muted-foreground">Layout “{t.layout}” · {t.slots.length} slots</span>
                </div>
              );
            })}
          </div>
        )}
      </Section>
    </ScrollPage>
  );
}
