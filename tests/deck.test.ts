import { describe, expect, test } from "bun:test";
import { Deck } from "@/features/presenter/pptx/deck";
import { slideTypes, typeOf } from "@/features/presenter/pptx/template";
import { sampleDeck, templateDeck } from "./fixtures";

const reopen = async (d: Deck) => Deck.open(await d.bytes());
const PNG = new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64"));

describe("pptx deck", () => {
  test("reads the slides as text: shapes, places, room, notes", async () => {
    const d = await Deck.open(await sampleDeck());
    expect(await d.count()).toBe(3);
    const s = await d.slide(2);
    expect(s.title).toBe("Results Q2");
    expect(s.notes).toContain("finance model");
    const body = s.shapes.find((x) => x.name === "Body")!;
    expect(body.paragraphs.map((p) => p.text)).toEqual(["Revenue 11.2 M€", "Margin 36 %"]);
    expect(body.budget).toBeGreaterThan(20);
    expect(s.shapes.some((x) => x.kind === "picture" && x.image)).toBe(true);
    expect(await d.outline()).toContain('"Body"');
  });

  test("sets a text keeping the shape's look, and replaces across the deck", async () => {
    const d = await Deck.open(await sampleDeck());
    const body = (await d.slide(2)).shapes.find((x) => x.name === "Body")!;
    await d.setText(2, body.id, "Revenue 12.4 M€\n  up 8 %\nMargin 38 %");
    expect(await d.replaceText("Q2", "Q3")).toBe(2);
    const again = await reopen(d);
    const s = await again.slide(2);
    const p = s.shapes.find((x) => x.name === "Body")!.paragraphs;
    expect(p.map((x) => x.text)).toEqual(["Revenue 12.4 M€", "up 8 %", "Margin 38 %"]);
    expect(p[1].level).toBe(1);
    expect(p[0].size).toBe(18);
    expect(s.title).toBe("Results Q3");
    expect((await again.slide(1)).title).toBe("Quarterly review Q3");
  });

  test("adds a slide of the template, moves, duplicates and deletes slides", async () => {
    const d = await Deck.open(await sampleDeck());
    const t = await Deck.open(await templateDeck());
    const types = await slideTypes(t);
    expect(types.map((x) => x.name)).toEqual(["Two columns", "Visual"]);
    const visual = typeOf(types, "visual")!;
    await d.copySlide(t, visual.number, 3);
    const added = await d.slide(3);
    expect(added.title).toBe("Key figure");
    expect(added.shapes.some((x) => x.kind === "picture")).toBe(true);
    await d.setText(3, added.shapes.find((x) => x.name === "Title")!.id, "Margin by region");
    await d.copySlide(d, 1, 2);
    await d.moveSlide(5, 1);
    await d.deleteSlide(2);
    const again = await reopen(d);
    expect((await again.slides()).map((s) => s.title)).toEqual(["Thank you", "Quarterly review Q2", "Results Q2", "Margin by region"]);
  });

  test("puts a picture in place of a shape and at a box", async () => {
    const d = await Deck.open(await sampleDeck());
    const body = (await d.slide(2)).shapes.find((x) => x.name === "Body")!;
    await d.putPicture(2, PNG, { width: 400, height: 300 }, { replace: body.id });
    await d.putPicture(3, PNG, { width: 800, height: 450 }, { box: { x: 10, y: 10, w: 80, h: 60 } });
    const again = await reopen(d);
    const s2 = await again.slide(2);
    expect(s2.shapes.some((x) => x.name === "Body")).toBe(false);
    expect(s2.shapes.filter((x) => x.kind === "picture").length).toBe(2);
    expect((await again.slide(3)).shapes.some((x) => x.kind === "picture" && x.image?.startsWith("ppt/media/hub-image"))).toBe(true);
  });

  test("raw XML: read, refused when invalid, written", async () => {
    const d = await Deck.open(await sampleDeck());
    const xml = await d.xml(3);
    expect(xml).toContain("Thank you");
    await expect(d.setXml(3, "<p:sld")).rejects.toThrow();
    await d.setXml(3, xml.replace("Thank you", "Questions?"));
    expect((await (await reopen(d)).slide(3)).title).toBe("Questions?");
  });

  test("texts keep bullets and inline bold through a rewrite", async () => {
    const d = await Deck.open(await sampleDeck());
    const body = (await d.slide(2)).shapes.find((x) => x.name === "Body")!;
    await d.setText(2, body.id, "Intro line\n- Revenue **12.4 M€**\n- Margin 38 %\n  - up 2 pts");
    const read = (await (await reopen(d)).slide(2)).shapes.find((x) => x.name === "Body")!.paragraphs;
    expect(read.map((p) => `${"  ".repeat(p.level)}${p.bullet ? "- " : ""}${p.marked}`)).toEqual(["Intro line", "- Revenue **12.4 M€**", "- Margin 38 %", "  - up 2 pts"]);
    // Written back as read: the same.
    await d.setText(2, body.id, read.map((p) => `${"  ".repeat(p.level)}${p.bullet ? "- " : ""}${p.marked}`).join("\n"));
    expect((await d.slide(2)).shapes.find((x) => x.name === "Body")!.paragraphs.map((p) => p.marked)).toEqual(read.map((p) => p.marked));
  });

  test("a new text box is placed by code and takes the slide's look", async () => {
    const d = await Deck.open(await sampleDeck());
    const before = await d.slide(2);
    const body = before.shapes.find((x) => x.name === "Body")!;
    const id = await d.addText(2, "Location types:\n- Store\n- Warehouse", { below: body.id });
    const added = (await (await reopen(d)).slide(2)).shapes.find((x) => x.id === id)!;
    expect(added.kind).toBe("text");
    expect(added.box!.x).toBe(body.box!.x);
    expect(added.box!.y).toBeGreaterThan(body.box!.y + body.box!.h);
    expect(added.size).toBe(body.size);
    expect(added.paragraphs.map((p) => p.bullet)).toEqual([false, true, true]);
    // Without a place: the largest free area, touching no shape with content.
    const free = await d.placeFor(3);
    const closing = (await d.slide(3)).shapes[0].box!;
    expect(free.y >= closing.y + closing.h || free.y + free.h <= closing.y).toBe(true);
    await expect(d.addText(2, "x", { below: "#999" })).rejects.toThrow(/No shape #999/);
  });

  test("tables: added in the slide's look, rows replaced, added and removed", async () => {
    const d = await Deck.open(await sampleDeck());
    const id = await d.addTable(3, [["Type", "Stores"], ["Store", "120"], ["Warehouse", "8"]], { below: "Closing" });
    await d.setTable(3, id, [["Type", "Stores"], ["Store", "121"], ["Warehouse", "8"], ["Others", "3"]]);
    const t = (await (await reopen(d)).slide(3)).shapes.find((x) => x.id === id)!;
    expect(t.kind).toBe("table");
    expect(t.rows).toEqual([["Type", "Stores"], ["Store", "121"], ["Warehouse", "8"], ["Others", "3"]]);
    await expect(d.setTable(3, id, [["a", "b", "c"]])).rejects.toThrow(/2 columns/);
    await expect(d.setText(3, id, "x")).rejects.toThrow(/set_table/);
  });

  test("frames and decoration are told apart, and the slide reads in order", async () => {
    const d = await Deck.open(await sampleDeck());
    const body = (await d.slide(2)).shapes.find((x) => x.name === "Body")!;
    const frame = await d.addBox(2, { box: { x: 2, y: 18, w: 50, h: 50 } });
    const dot = await d.addBox(2, { box: { x: 90, y: 90, w: 2, h: 2 } }, { fill: "#112233" });
    const s = await d.slide(2);
    expect(s.shapes.find((x) => x.id === frame)!.role).toBe("frame");
    expect(s.shapes.find((x) => x.id === dot)!.role).toBe("decoration");
    expect(s.shapes.find((x) => x.id === body.id)!.frame).toBe(frame);
    const text = d.describe(s);
    expect(text).toContain(`frame #${frame}`);
    expect(text).toContain("1 decorative shape(s) without text not listed");
    expect(text).not.toContain(`#${dot} `);
    expect(text.indexOf('"Heading"')).toBeLessThan(text.indexOf('"Body"'));
    // A text added inside the frame avoids what is already in it.
    const id = await d.addText(2, "Note", { inside: frame });
    const note = (await d.slide(2)).shapes.find((x) => x.id === id)!.box!;
    const b = body.box!;
    expect(note.x >= b.x + b.w || note.x + note.w <= b.x || note.y >= b.y + b.h || note.y + note.h <= b.y).toBe(true);
  });

  test("a slide from a layout of the deck, its placeholders filled by role; the look of a shape changed", async () => {
    const d = await Deck.open(await sampleDeck());
    const layouts = await d.layouts();
    expect(layouts.length).toBeGreaterThan(0);
    const name = await d.addSlideFromLayout(undefined, 2);
    expect(layouts.map((l) => l.name)).toContain(name);
    expect(await d.count()).toBe(4);
    const heading = (await d.slide(3)).shapes.find((x) => x.name === "Heading")!;
    await d.setStyle(3, `#${heading.id}`, { size: 28, bold: true, color: "#112233", fill: "#EEEEEE" });
    const again = (await (await reopen(d)).slide(3)).shapes.find((x) => x.name === "Heading")!;
    expect(again.size).toBe(28);
    expect(again.paragraphs[0].color).toBe("#112233");
    expect(again.fill).toBe("#EEEEEE");
    expect(await d.fonts()).toBeArray();
  });
});

