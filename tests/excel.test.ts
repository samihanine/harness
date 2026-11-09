import { describe, expect, test } from "bun:test";
import { checkValue, guessColumns, readData, reconcile, rowCells, rowMatches, viewRows, type Column } from "@/features/excel/table";

const COLUMNS: Column[] = [
  { name: "Title", type: "text" },
  { name: "Status", type: "option", options: [{ value: "Open", color: "blue" }, { value: "Done", color: "green" }] },
  { name: "Due", type: "date" },
];
const HEADER = ["Title", "Status", "Due", "Owner"];
const CELLS = [
  ["Refresh the model", "Open", 46335, "Sami"],
  ["Write the guide", "Done", "2026-10-01", "Alex"],
  ["Review the deck", "Open", "", "Sami"],
];

describe("excel table", () => {
  test("reads dates (Excel numbers too) and keeps unknown columns", () => {
    const d = readData(COLUMNS, HEADER, CELLS);
    expect(d.rows[0]).toEqual({ Title: "Refresh the model", Status: "Open", Due: "2026-11-09", Owner: "Sami" });
    expect(d.rows[1].Due).toBe("2026-10-01");
  });

  test("a row's cells keep the file's values not given", () => {
    expect(rowCells(COLUMNS, HEADER, CELLS[0], { Status: "Done" })).toEqual(["Refresh the model", "Done", 46335, "Sami"]);
    expect(rowCells(COLUMNS, HEADER, undefined, { Title: "New" })).toEqual(["New", "", "", ""]);
  });

  test("values are checked against their column", () => {
    expect(checkValue(COLUMNS[1], "done")).toBe("Done");
    expect(() => checkValue(COLUMNS[1], "Late")).toThrow(/not an option/);
    expect(checkValue(COLUMNS[2], "2026-12-01")).toBe("2026-12-01");
    expect(() => checkValue(COLUMNS[2], "someday")).toThrow(/not a date/);
  });

  test("conditions select rows without ids", () => {
    const { rows } = readData(COLUMNS, HEADER, CELLS);
    const pick = (where: Parameters<typeof rowMatches>[1]) => rows.filter((r) => rowMatches(r, where, COLUMNS)).map((r) => r.Title);
    expect(pick({ status: "open" })).toEqual(["Refresh the model", "Review the deck"]);
    expect(pick({ Due: { empty: true } })).toEqual(["Review the deck"]);
    expect(pick({ Due: { before: "2026-11-01" } })).toEqual(["Write the guide"]);
    expect(pick({ Title: { contains: "the d" } })).toEqual(["Review the deck"]);
    expect(() => pick({ Priority: "High" })).toThrow(/No column/);
  });

  test("views filter, search and sort", () => {
    const { rows } = readData(COLUMNS, HEADER, CELLS);
    const shown = viewRows(rows, { id: "v", name: "v", layout: "table", filters: [{ column: "Status", values: ["Open"] }], sort: { column: "Title", desc: true } });
    expect(shown.map((r) => r.index)).toEqual([2, 0]);
  });

  test("columns: new ones of the file added, the app's pending ones listed", () => {
    const r = reconcile([...COLUMNS, { name: "Priority", type: "text" }], HEADER, CELLS);
    expect(r.columns.map((c) => c.name)).toEqual(["Title", "Status", "Due", "Priority", "Owner"]);
    expect(r.columns.at(-1)).toMatchObject({ name: "Owner", type: "option" });
    expect(guessColumns(["S"], [["a"], ["b"], ["a"], ["b"]])[0]).toMatchObject({ type: "option", options: [{ value: "a" }, { value: "b" }] });
    expect(r.pending.add).toEqual(["Priority"]);
    const renamed = reconcile([{ ...COLUMNS[0], name: "Task", was: "Title" }, ...COLUMNS.slice(1)], HEADER, CELLS);
    expect(renamed.pending.reorder).toBe(true);
    expect(reconcile(guessColumns(HEADER, CELLS), HEADER, CELLS).dirty).toBe(false);
  });
});

describe("xlsx choices", () => {
  test("option columns get a drop-down list and a fill per option, in schema order, idempotent", async () => {
    const { OoxmlPackage } = await import("@/lib/ooxml/package");
    const { applyChoices } = await import("@/features/excel/xlsx");
    const { sampleWorkbook } = await import("./fixtures");
    const pkg = await OoxmlPackage.open(await sampleWorkbook());
    const columns = COLUMNS.slice(0, 3);
    await applyChoices(pkg, "Tasks", columns, ["Title", "Status", "Due"]);
    await applyChoices(pkg, "Tasks", columns, ["Title", "Status", "Due"]);
    const sheet = await pkg.text("xl/worksheets/sheet1.xml");
    expect(sheet.match(/<dataValidation /g)?.length).toBe(1);
    expect(sheet).toContain('sqref="B2:B1048576"');
    expect(sheet).toContain("<formula1>\"Open,Done\"</formula1>");
    expect(sheet.match(/<conditionalFormatting /g)?.length).toBe(1);
    expect(sheet.match(/<cfRule /g)?.length).toBe(2);
    // conditionalFormatting, then dataValidations, before pageMargins (the schema's order).
    expect(sheet.indexOf("<conditionalFormatting")).toBeLessThan(sheet.indexOf("<dataValidations"));
    expect(sheet.indexOf("<dataValidations")).toBeLessThan(sheet.indexOf("<pageMargins"));
    const styles = await pkg.text("xl/styles.xml");
    expect(styles.match(/<dxf>/g)?.length).toBe(2);
    expect(styles.indexOf("<dxfs")).toBeLessThan(styles.indexOf("<tableStyles"));
    await expect(applyChoices(pkg, "Tasks", columns, ["Title", "Due", "Status"])).rejects.toThrow("NOT_READY");
  });
});
