import { describe, expect, it } from "vitest";
import type { Row } from "../src/row";
import { Workbook } from "../src/workbook";
import { type SheetInput, xlsx } from "./support/xlsx-fixture";

async function rowsOf(sheet: SheetInput): Promise<Row[]> {
  const workbook = await Workbook.open(xlsx([sheet]));
  const rows: Row[] = [];
  for await (const row of workbook.worksheet(sheet.name).rows()) {
    rows.push(row);
  }
  return rows;
}

describe("Worksheet rows", () => {
  it("reads number and shared-string cells", async () => {
    const rows = await rowsOf({
      name: "Data",
      rows: [
        [1, "Hello"],
        ["World", 2.5],
      ],
    });

    expect(rows).toEqual([
      {
        number: 1,
        cells: [
          { ref: "A1", columnIndex: 0, type: "number", value: 1 },
          { ref: "B1", columnIndex: 1, type: "string", value: "Hello" },
        ],
      },
      {
        number: 2,
        cells: [
          { ref: "A2", columnIndex: 0, type: "string", value: "World" },
          { ref: "B2", columnIndex: 1, type: "number", value: 2.5 },
        ],
      },
    ]);
  });

  it("reads boolean, error, inline string and formula string cells", async () => {
    const [row] = await rowsOf({
      name: "Data",
      rows: [
        [
          { boolean: true },
          { boolean: false },
          { error: "#DIV/0!" },
          { inlineString: "inline" },
          { formulaString: "formula" },
        ],
      ],
    });

    expect(row?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "boolean", value: true },
      { ref: "B1", columnIndex: 1, type: "boolean", value: false },
      { ref: "C1", columnIndex: 2, type: "error", value: "#DIV/0!" },
      { ref: "D1", columnIndex: 3, type: "string", value: "inline" },
      { ref: "E1", columnIndex: 4, type: "string", value: "formula" },
    ]);
  });

  it("reads a formula cell as its text and cached value", async () => {
    const [row] = await rowsOf({
      name: "Data",
      rows: [[{ formula: "B1+C1", cached: 5 }, 2, 3]],
    });

    expect(row?.cells[0]).toEqual({
      ref: "A1",
      columnIndex: 0,
      type: "formula",
      value: "B1+C1",
      cachedValue: { type: "number", value: 5 },
    });
  });

  it("keeps a formula cell that has no cached value", async () => {
    const [row] = await rowsOf({
      name: "Data",
      rows: [[{ formula: "SUM(B1:B3)" }, "after"]],
    });

    expect(row?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "formula", value: "SUM(B1:B3)", cachedValue: null },
      { ref: "B1", columnIndex: 1, type: "string", value: "after" },
    ]);
  });

  it("reads a styled number as a date", async () => {
    const date = new Date(Date.UTC(2020, 0, 15));

    const [row] = await rowsOf({ name: "Data", rows: [[date]] });

    expect(row?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "date", value: date }]);
  });

  it("reads dates under the 1904 date system", async () => {
    const date = new Date(Date.UTC(2020, 0, 15));
    const workbook = await Workbook.open(xlsx([{ name: "Data", rows: [[date]] }], { date1904: true }));

    const rows: Row[] = [];
    for await (const row of workbook.worksheet("Data").rows()) {
      rows.push(row);
    }

    expect(rows[0]?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "date", value: date }]);
  });

  it("reads a styled number with a time of day as a date and time", async () => {
    const date = new Date(Date.UTC(2020, 0, 15, 13, 30));

    const [row] = await rowsOf({ name: "Data", rows: [[date]] });

    expect(row?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "date", value: date }]);
  });

  it("reads an explicit t=d date cell from its ISO value", async () => {
    const [row] = await rowsOf({ name: "Data", rows: [[{ isoDate: "2020-01-15T13:30:00" }]] });

    expect(row?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "date", value: new Date(Date.UTC(2020, 0, 15, 13, 30)) },
    ]);
  });

  it("throws on an unsupported cell type", async () => {
    await expect(rowsOf({ name: "Data", rows: [[{ rawType: "zzz" }]] })).rejects.toThrow(
      'Unsupported cell type "zzz" at A1',
    );
  });
});
