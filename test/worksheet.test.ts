import { describe, expect, it } from "vitest";
import type { Row } from "../src/xlsx/cell";
import { Workbook } from "../src/xlsx/workbook";
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
          { ref: "A1", type: "number", value: 1 },
          { ref: "B1", type: "string", value: "Hello" },
        ],
      },
      {
        number: 2,
        cells: [
          { ref: "A2", type: "string", value: "World" },
          { ref: "B2", type: "number", value: 2.5 },
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
      { ref: "A1", type: "boolean", value: true },
      { ref: "B1", type: "boolean", value: false },
      { ref: "C1", type: "error", value: "#DIV/0!" },
      { ref: "D1", type: "string", value: "inline" },
      { ref: "E1", type: "string", value: "formula" },
    ]);
  });

  it("reads a styled number as a date", async () => {
    const date = new Date(Date.UTC(2020, 0, 15));

    const [row] = await rowsOf({ name: "Data", rows: [[date]] });

    expect(row?.cells).toEqual([{ ref: "A1", type: "date", value: date }]);
  });

  it("reads dates under the 1904 date system", async () => {
    const date = new Date(Date.UTC(2020, 0, 15));
    const workbook = await Workbook.open(xlsx([{ name: "Data", rows: [[date]] }], { date1904: true }));

    const rows: Row[] = [];
    for await (const row of workbook.worksheet("Data").rows()) {
      rows.push(row);
    }

    expect(rows[0]?.cells).toEqual([{ ref: "A1", type: "date", value: date }]);
  });

  it("throws on an unsupported cell type", async () => {
    await expect(rowsOf({ name: "Data", rows: [[{ rawType: "d" }]] })).rejects.toThrow(
      'Unsupported cell type "d" at A1',
    );
  });
});
