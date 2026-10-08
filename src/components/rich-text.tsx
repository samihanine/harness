/**
 * Rich text editor (Tiptap) whose value is markdown: what is stored in the Excel cell stays readable
 * and is shown as is by Power BI. Shortcuts: ⌘B, ⌘I, ⌘⇧X, "# " titles, "- " lists, "> " quotes…
 */
import { useEffect } from "react";
import { EditorContent, useEditor, useEditorState, type Editor } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import { Markdown } from "@tiptap/markdown";
import { BoldIcon, CodeIcon, Heading2Icon, Heading3Icon, ItalicIcon, LinkIcon, ListIcon, ListOrderedIcon, QuoteIcon, StrikethroughIcon, UnderlineIcon } from "lucide-react";

export function RichText({ value, onChange, autoFocus }: { value: string; onChange: (markdown: string) => void; autoFocus?: boolean }) {
  const editor = useEditor({
    extensions: [StarterKit.configure({ link: { openOnClick: false, autolink: true } }), Markdown],
    content: value,
    contentType: "markdown",
    immediatelyRender: false,
    autofocus: autoFocus ? "end" : false,
    editorProps: { attributes: { class: "markdown min-h-24 px-3 py-2 outline-none" } },
    onBlur: ({ editor }) => {
      const next = editor.getMarkdown();
      if (next !== value) onChange(next);
    },
  });
  // Value changed elsewhere (AI, reload): shown unless the user is typing.
  useEffect(() => {
    if (editor && !editor.isFocused && editor.getMarkdown() !== value) editor.commands.setContent(value, { contentType: "markdown" });
  }, [editor, value]);
  return (
    <div className="rounded-md border bg-card focus-within:border-ring">
      {editor && <Toolbar editor={editor} />}
      <EditorContent editor={editor} />
    </div>
  );
}

function Toolbar({ editor }: { editor: Editor }) {
  const active = useEditorState({
    editor,
    selector: ({ editor: e }) => ({
      bold: e.isActive("bold"),
      italic: e.isActive("italic"),
      underline: e.isActive("underline"),
      strike: e.isActive("strike"),
      code: e.isActive("code"),
      h2: e.isActive("heading", { level: 2 }),
      h3: e.isActive("heading", { level: 3 }),
      bullet: e.isActive("bulletList"),
      ordered: e.isActive("orderedList"),
      quote: e.isActive("blockquote"),
      link: e.isActive("link"),
    }),
  });
  const chain = () => editor.chain().focus();
  const link = () => {
    const url = prompt("Link", editor.getAttributes("link").href ?? "https://");
    if (url === null) return;
    if (url) chain().extendMarkRange("link").setLink({ href: url }).run();
    else chain().extendMarkRange("link").unsetLink().run();
  };
  const buttons = [
    { on: active.bold, icon: <BoldIcon />, title: "Bold (⌘B)", run: () => chain().toggleBold().run() },
    { on: active.italic, icon: <ItalicIcon />, title: "Italic (⌘I)", run: () => chain().toggleItalic().run() },
    { on: active.underline, icon: <UnderlineIcon />, title: "Underline (⌘U)", run: () => chain().toggleUnderline().run() },
    { on: active.strike, icon: <StrikethroughIcon />, title: "Strikethrough (⌘⇧X)", run: () => chain().toggleStrike().run() },
    { on: active.code, icon: <CodeIcon />, title: "Code (⌘E)", run: () => chain().toggleCode().run() },
    "|",
    { on: active.h2, icon: <Heading2Icon />, title: "Title (## )", run: () => chain().toggleHeading({ level: 2 }).run() },
    { on: active.h3, icon: <Heading3Icon />, title: "Subtitle (### )", run: () => chain().toggleHeading({ level: 3 }).run() },
    { on: active.bullet, icon: <ListIcon />, title: "List (- )", run: () => chain().toggleBulletList().run() },
    { on: active.ordered, icon: <ListOrderedIcon />, title: "Numbered list (1. )", run: () => chain().toggleOrderedList().run() },
    { on: active.quote, icon: <QuoteIcon />, title: "Quote (> )", run: () => chain().toggleBlockquote().run() },
    { on: active.link, icon: <LinkIcon />, title: "Link", run: link },
  ] as const;
  return (
    <div className="flex flex-wrap items-center gap-0.5 border-b px-1.5 py-1" onMouseDown={(e) => e.preventDefault()}>
      {buttons.map((b, i) =>
        b === "|" ? (
          <span key={i} className="mx-1 h-4 w-px bg-border" />
        ) : (
          <button key={i} type="button" title={b.title} onClick={b.run} className={`inline-flex size-7 items-center justify-center rounded text-muted-foreground hover:bg-muted hover:text-foreground [&_svg]:size-3.5 ${b.on ? "bg-muted text-foreground" : ""}`}>
            {b.icon}
          </button>
        ),
      )}
    </div>
  );
}
