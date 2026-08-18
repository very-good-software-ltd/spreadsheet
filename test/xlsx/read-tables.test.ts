import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { BytesByteRange } from "../../src/io/byte-range";
import { readTables } from "../../src/xlsx/read-tables";
import { createXmlReader } from "../../src/xml/create-xml-reader";
import { openZip } from "../../src/zip/open-zip";
import { type SheetInput, xlsx } from "../support/xlsx-fixture";

async function tablesOf(sheets: readonly SheetInput[], sheetNumber = 1) {
  const archive = await openZip(new BytesByteRange(xlsx(sheets)));
  return readTables(archive, createXmlReader(), `xl/worksheets/sheet${sheetNumber}.xml`);
}

const ROWS = [[1]];

describe("readTables", () => {
  it("reads a table's name and the extent it covers", async () => {
    const tables = await tablesOf([{ name: "Sheet1", rows: ROWS, tables: [{ name: "Sales", ref: "B2:D10" }] }]);

    expect(tables).toEqual([
      {
        name: "Sales",
        path: "xl/tables/table1.xml",
        firstRow: 2,
        lastRow: 10,
        firstColumnIndex: 1,
        lastColumnIndex: 3,
        headerRowCount: 1,
        totalsRowCount: 0,
      },
    ]);
  });

  // A table carries both, and displayName is the one Excel shows and the one a
  // formula refers to. They agree in files Excel writes, but the spec lets them
  // differ, and the name a caller would type is the display one.
  it("takes the name Excel shows when the two disagree", async () => {
    const tables = await tablesOf([
      { name: "Sheet1", rows: ROWS, tables: [{ name: "Table1", displayName: "Sales", ref: "A1:B2" }] },
    ]);

    expect(tables[0]?.name).toBe("Sales");
  });

  it("reads a header row and a totals row when the table declares them", async () => {
    const tables = await tablesOf([
      {
        name: "Sheet1",
        rows: ROWS,
        tables: [{ name: "Sales", ref: "A1:B9", headerRowCount: 1, totalsRowCount: 1 }],
      },
    ]);

    expect(tables[0]).toMatchObject({ headerRowCount: 1, totalsRowCount: 1 });
  });

  // The attribute is optional and its default is one, so a table without it still
  // has a header row. Reading it as zero would treat the headings as data.
  it("counts one header row when the table does not say", async () => {
    const tables = await tablesOf([{ name: "Sheet1", rows: ROWS, tables: [{ name: "Sales", ref: "A1:B9" }] }]);

    expect(tables[0]?.headerRowCount).toBe(1);
  });

  it("reads a table that declares it has no header row", async () => {
    const tables = await tablesOf([
      { name: "Sheet1", rows: ROWS, tables: [{ name: "Sales", ref: "A1:B9", headerRowCount: 0 }] },
    ]);

    expect(tables[0]?.headerRowCount).toBe(0);
  });

  it("finds every table a sheet has", async () => {
    const tables = await tablesOf([
      {
        name: "Sheet1",
        rows: ROWS,
        tables: [
          { name: "Sales", ref: "A1:B9" },
          { name: "Costs", ref: "D1:E9" },
        ],
      },
    ]);

    expect(tables.map((table) => table.name)).toEqual(["Sales", "Costs"]);
  });

  it("has none for a sheet with no relationships of its own", async () => {
    expect(await tablesOf([{ name: "Sheet1", rows: ROWS }])).toEqual([]);
  });

  it("keeps each sheet's tables to that sheet", async () => {
    const sheets: SheetInput[] = [
      { name: "First", rows: ROWS, tables: [{ name: "Sales", ref: "A1:B9" }] },
      { name: "Second", rows: ROWS, tables: [{ name: "Costs", ref: "A1:B9" }] },
    ];

    expect((await tablesOf(sheets, 2)).map((table) => table.name)).toEqual(["Costs"]);
  });
});

// The fixture above is our own encoding of the format, so it could just as easily
// encode a misunderstanding of it. exceljs writes a real table part, with the
// relative target that climbs out of the worksheets folder and the display name
// alongside the plain one.
describe("readTables against a table exceljs wrote", () => {
  it("reads its name and its extent, totals row included", async () => {
    const source = new ExcelJS.Workbook();
    source.addWorksheet("Sheet1").addTable({
      name: "Sales",
      ref: "B2",
      headerRow: true,
      totalsRow: true,
      columns: [
        { name: "Item", totalsRowLabel: "Total" },
        { name: "Amount", totalsRowFunction: "sum" },
      ],
      rows: [
        ["a", 1],
        ["b", 2],
      ],
    });

    const archive = await openZip(new BytesByteRange(new Uint8Array(await source.xlsx.writeBuffer())));
    const tables = await readTables(archive, createXmlReader(), "xl/worksheets/sheet1.xml");

    expect(tables).toEqual([
      {
        name: "Sales",
        path: "xl/tables/table1.xml",
        firstRow: 2,
        lastRow: 5,
        firstColumnIndex: 1,
        lastColumnIndex: 2,
        headerRowCount: 1,
        totalsRowCount: 1,
      },
    ]);
  });
});
