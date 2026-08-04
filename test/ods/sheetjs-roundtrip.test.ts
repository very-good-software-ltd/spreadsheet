import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import type { Row } from "../../src/row";
import { Workbook } from "../../src/workbook";

// SheetJS writes the .ods and we read it back, so a separate implementation
// produces the input rather than our own fixture encoder. It is not a full
// stand-in for real ODF, though: it writes a date as a plain float, the serial
// day count, rather than an ODF date cell, and its output is compact, so it
// never emits the run-length repeats. Dates and repeats are covered by
// read-ods.test.ts against real ODF markup instead.
async function readOds(rows: (string | number | boolean)[][], sheetName: string): Promise<Row[]> {
  const source = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(source, XLSX.utils.aoa_to_sheet(rows), sheetName);
  const bytes = new Uint8Array(XLSX.write(source, { bookType: "ods", type: "buffer" }));

  const workbook = await Workbook.open(bytes);
  const read: Row[] = [];
  for await (const row of workbook.worksheet(sheetName).rows()) {
    read.push(row);
  }
  return read;
}

describe("reading SheetJS-written ods", () => {
  it("reads string and number cells", async () => {
    const rows = await readOds(
      [
        ["SKU", "Weight"],
        ["PROD-001", 150],
      ],
      "Data",
    );

    expect(rows).toEqual([
      {
        number: 1,
        cells: [
          { ref: "A1", columnIndex: 0, type: "string", value: "SKU" },
          { ref: "B1", columnIndex: 1, type: "string", value: "Weight" },
        ],
      },
      {
        number: 2,
        cells: [
          { ref: "A2", columnIndex: 0, type: "string", value: "PROD-001" },
          { ref: "B2", columnIndex: 1, type: "number", value: 150 },
        ],
      },
    ]);
  });

  it("reads boolean cells", async () => {
    const rows = await readOds([[true, false]], "Data");

    expect(rows[0]?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "boolean", value: true },
      { ref: "B1", columnIndex: 1, type: "boolean", value: false },
    ]);
  });

  it("reads a formula cell with its cached result", async () => {
    const source = XLSX.utils.book_new();
    const sheet: XLSX.WorkSheet = { "!ref": "A1:A1", A1: { t: "n", f: "2*5", v: 10 } };
    XLSX.utils.book_append_sheet(source, sheet, "Data");
    const bytes = new Uint8Array(XLSX.write(source, { bookType: "ods", type: "buffer" }));

    const workbook = await Workbook.open(bytes);
    const rows: Row[] = [];
    for await (const row of workbook.worksheet("Data").rows()) {
      rows.push(row);
    }

    expect(rows[0]?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "formula", value: "of:=2*5", cachedValue: { type: "number", value: 10 } },
    ]);
  });
});
