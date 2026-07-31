import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import type { Row } from "../src/xlsx/row";
import { Workbook } from "../src/xlsx/workbook";

// exceljs writes the file and we read it back, so a separate implementation
// produces the input rather than our own fixture encoder.
async function readWith(build: (workbook: ExcelJS.Workbook) => void, sheetName: string): Promise<Row[]> {
  const source = new ExcelJS.Workbook();
  build(source);

  const workbook = await Workbook.open(await source.xlsx.writeBuffer());
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
    const rows = await readWith((workbook) => {
      workbook.addWorksheet("Data").addRow([true, false]);
    }, "Data");

    expect(rows[0]?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "boolean", value: true },
      { ref: "B1", columnIndex: 1, type: "boolean", value: false },
    ]);
  });

  it("looks a cell up by column, returning undefined for a gap", async () => {
    const rows = await readWith((workbook) => {
      const sheet = workbook.addWorksheet("Data");
      sheet.getCell("A1").value = "left";
      sheet.getCell("C1").value = "right";
    }, "Data");
    const row = rows[0];

    expect(row?.cell(0)).toEqual({ ref: "A1", columnIndex: 0, type: "string", value: "left" });
    expect(row?.cell(2)).toEqual({ ref: "C1", columnIndex: 2, type: "string", value: "right" });
    expect(row?.cell(1)).toBeUndefined();
  });

  it("reads a date cell written with a date number format", async () => {
    const date = new Date(Date.UTC(2020, 0, 15));

    const rows = await readWith((workbook) => {
      const cell = workbook.addWorksheet("Data").getCell("A1");
      cell.value = date;
      cell.numFmt = "yyyy-mm-dd";
    }, "Data");

    expect(rows[0]?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "date", value: date }]);
  });

  it("reads a formula cell as its cached result", async () => {
    const rows = await readWith((workbook) => {
      workbook.addWorksheet("Data").getCell("A1").value = { formula: "2*5", result: 10 };
    }, "Data");

    expect(rows[0]?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "number", value: 10 }]);
  });

  it("reports a hidden sheet as hidden", async () => {
    const source = new ExcelJS.Workbook();
    source.addWorksheet("Visible").addRow(["x"]);
    const hidden = source.addWorksheet("Hidden");
    hidden.addRow(["y"]);
    hidden.state = "hidden";

    const workbook = await Workbook.open(await source.xlsx.writeBuffer());

    expect(workbook.worksheets).toEqual([
      { name: "Visible", hidden: false },
      { name: "Hidden", hidden: true },
    ]);
  });
});
