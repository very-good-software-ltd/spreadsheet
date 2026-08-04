import { describe, expect, it } from "vitest";
import type { Row } from "../../src/row";
import { Workbook } from "../../src/workbook";
import { odsWith } from "../support/ods-fixture";

async function rowsOf(bytes: Uint8Array, sheet: string): Promise<Row[]> {
  const workbook = await Workbook.open(bytes);
  const rows: Row[] = [];
  for await (const row of workbook.worksheet(sheet).rows()) {
    rows.push(row);
  }
  return rows;
}

describe("readOds", () => {
  it("reads a cell and skips the empty column padding", async () => {
    const [row] = await rowsOf(
      odsWith(
        `<table:table table:name="S"><table:table-row>` +
          `<table:table-cell office:value-type="float" office:value="42"><text:p>42</text:p></table:table-cell>` +
          `<table:table-cell table:number-columns-repeated="16383"/>` +
          `</table:table-row></table:table>`,
      ),
      "S",
    );

    expect(row?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "number", value: 42 }]);
  });

  it("advances the column index across an empty gap", async () => {
    const [row] = await rowsOf(
      odsWith(
        `<table:table table:name="S"><table:table-row>` +
          `<table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell>` +
          `<table:table-cell table:number-columns-repeated="2"/>` +
          `<table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell>` +
          `</table:table-row></table:table>`,
      ),
      "S",
    );

    expect(row?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "string", value: "a" },
      { ref: "D1", columnIndex: 3, type: "string", value: "b" },
    ]);
  });

  it("expands a repeated non-empty cell", async () => {
    const [row] = await rowsOf(
      odsWith(
        `<table:table table:name="S"><table:table-row>` +
          `<table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="3"><text:p>7</text:p></table:table-cell>` +
          `</table:table-row></table:table>`,
      ),
      "S",
    );

    expect(row?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "number", value: 7 },
      { ref: "B1", columnIndex: 1, type: "number", value: 7 },
      { ref: "C1", columnIndex: 2, type: "number", value: 7 },
    ]);
  });

  it("skips empty rows and numbers the rows past the gap", async () => {
    const rows = await rowsOf(
      odsWith(
        `<table:table table:name="S">` +
          `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>` +
          `<table:table-row table:number-rows-repeated="5"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>` +
          `<table:table-row><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell></table:table-row>` +
          `</table:table>`,
      ),
      "S",
    );

    expect(rows.map((row) => [row.number, row.cells[0]?.value])).toEqual([
      [1, "a"],
      [7, "b"],
    ]);
  });

  it("reads boolean, date, and formula cells", async () => {
    const [row] = await rowsOf(
      odsWith(
        `<table:table table:name="S"><table:table-row>` +
          `<table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>` +
          `<table:table-cell office:value-type="date" office:date-value="2020-01-15"><text:p>2020-01-15</text:p></table:table-cell>` +
          `<table:table-cell office:value-type="float" office:value="10" table:formula="of:=5+5"><text:p>10</text:p></table:table-cell>` +
          `</table:table-row></table:table>`,
      ),
      "S",
    );

    expect(row?.cells).toEqual([
      { ref: "A1", columnIndex: 0, type: "boolean", value: true },
      { ref: "B1", columnIndex: 1, type: "date", value: new Date(Date.UTC(2020, 0, 15)) },
      { ref: "C1", columnIndex: 2, type: "formula", value: "of:=5+5", cachedValue: { type: "number", value: 10 } },
    ]);
  });

  it("reports a hidden sheet via table:visibility", async () => {
    const workbook = await Workbook.open(
      odsWith(`<table:table table:name="Visible"/><table:table table:name="Hidden" table:visibility="collapse"/>`),
    );

    expect(workbook.worksheets).toEqual([
      { name: "Visible", hidden: false },
      { name: "Hidden", hidden: true },
    ]);
  });
});
