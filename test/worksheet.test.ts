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

  it("throws on an unsupported cell type", async () => {
    await expect(rowsOf({ name: "Data", rows: [[{ date: "2020-01-01" }]] })).rejects.toThrow(
      'Unsupported cell type "d" at A1',
    );
  });
});
