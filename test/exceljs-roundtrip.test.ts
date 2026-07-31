import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import type { Row } from "../src/xlsx/cell";
import { Workbook } from "../src/xlsx/workbook";

// exceljs writes the file and we read it back, so a separate implementation
// produces the input rather than our own fixture encoder.
async function readWith(build: (workbook: ExcelJS.Workbook) => void, sheetName: string): Promise<Row[]> {
  const source = new ExcelJS.Workbook();
  build(source);
  const bytes = new Uint8Array((await source.xlsx.writeBuffer()) as ArrayBuffer);

  const workbook = await Workbook.open(bytes);
  const rows: Row[] = [];
  for await (const row of workbook.worksheet(sheetName).rows()) {
    rows.push(row);
  }
  return rows;
}

describe("reading exceljs-written workbooks", () => {
  it("reads string and number cells", async () => {
    const rows = await readWith((workbook) => {
      const sheet = workbook.addWorksheet("Data");
      sheet.addRow(["SKU", "Weight"]);
      sheet.addRow(["PROD-001", 150]);
    }, "Data");

    expect(rows).toEqual([
      {
        number: 1,
        cells: [
          { ref: "A1", type: "string", value: "SKU" },
          { ref: "B1", type: "string", value: "Weight" },
        ],
      },
      {
        number: 2,
        cells: [
          { ref: "A2", type: "string", value: "PROD-001" },
          { ref: "B2", type: "number", value: 150 },
        ],
      },
    ]);
  });

  it("reads boolean cells", async () => {
    const rows = await readWith((workbook) => {
      workbook.addWorksheet("Data").addRow([true, false]);
    }, "Data");

    expect(rows[0]?.cells).toEqual([
      { ref: "A1", type: "boolean", value: true },
      { ref: "B1", type: "boolean", value: false },
    ]);
  });

  it("reads a date cell written with a date number format", async () => {
    const date = new Date(Date.UTC(2020, 0, 15));

    const rows = await readWith((workbook) => {
      const cell = workbook.addWorksheet("Data").getCell("A1");
      cell.value = date;
      cell.numFmt = "yyyy-mm-dd";
    }, "Data");

    expect(rows[0]?.cells).toEqual([{ ref: "A1", type: "date", value: date }]);
  });

  it("reads a formula cell as its cached result", async () => {
    const rows = await readWith((workbook) => {
      workbook.addWorksheet("Data").getCell("A1").value = { formula: "2*5", result: 10 };
    }, "Data");

    expect(rows[0]?.cells).toEqual([{ ref: "A1", type: "number", value: 10 }]);
  });
});
