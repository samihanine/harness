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
    await d.putPicture(2, PNG, { width: 400, height: 300 }, { shape: body.id });
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
});
