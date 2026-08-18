import { describe, expect, it } from "vitest";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { shiftSheetRows } from "../../src/xlsx/shift-sheet";
import type { DateStyles, SheetWritePlan } from "../../src/xlsx/write-sheet";
import { writeSheetPart } from "../../src/xlsx/write-sheet";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";

const SHEET_OPEN = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;
const dateStyles: DateStyles = { forDate: () => "99" };

function streamOf(text: string): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(new TextEncoder().encode(text));
      controller.close();
    },
  });
}

const plan: SheetWritePlan = {
  positioned: (async function* () {})(),
  appended: (async function* () {})(),
  inheritedRows: new Set(),
};

async function shifted(sheet: string, shift: RowShift): Promise<string> {
  const events = shiftSheetRows(new SaxesXmlReader().read(streamOf(sheet)), shift);
  let out = "";
  for await (const chunk of writeSheetPart(events, plan, { dateStyles, date1904: false })) {
    out += chunk;
  }
  return out;
}

// Two rows appear at row 11, the last row of a region covering rows 9 to 11.
const INSERT: RowShift = { sheet: "Report", at: 11, by: 2 };

// Rows 10 and 11 go away, leaving row 9 as the whole region.
const REMOVE: RowShift = { sheet: "Report", at: 10, by: -2 };

describe("writeSheetPart moving rows", () => {
  it("renumbers the rows at and below the move, and their cells with them", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="9"><c r="A9"><v>1</v></c></row><row r="12"><c r="A12"><v>2</v></c></row></sheetData></worksheet>`;

    const out = await shifted(sheet, INSERT);

    expect(out).toContain(`<row r="9"><c r="A9"><v>1</v></c></row>`);
    expect(out).toContain(`<row r="14"><c r="A14"><v>2</v></c></row>`);
  });

  it("drops the rows that went away", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="9"><c r="A9"><v>1</v></c></row><row r="10"><c r="A10"><v>2</v></c></row><row r="12"><c r="A12"><v>3</v></c></row></sheetData></worksheet>`;

    const out = await shifted(sheet, REMOVE);

    expect(out).not.toContain(`<v>2</v>`);
    expect(out).toContain(`<row r="10"><c r="A10"><v>3</v></c></row>`);
  });

  it("moves the references in a formula, wherever the formula itself sits", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="12"><c r="C12"><f>SUM(C9:C11)</f><v>0</v></c></row></sheetData></worksheet>`;

    const out = await shifted(sheet, INSERT);

    expect(out).toContain(`<c r="C14"><f>SUM(C9:C13)</f>`);
  });

  it("moves the range a shared formula covers", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="12"><c r="C12"><f t="shared" ref="C12:C20" si="0">C11*2</f></c></row></sheetData></worksheet>`;

    const out = await shifted(sheet, INSERT);

    expect(out).toContain(`<f t="shared" ref="C14:C22" si="0">C13*2</f>`);
  });

  it("moves a merged range", async () => {
    const sheet = `${SHEET_OPEN}<sheetData/><mergeCells count="1"><mergeCell ref="A12:C12"/></mergeCells></worksheet>`;

    expect(await shifted(sheet, INSERT)).toContain(`<mergeCell ref="A14:C14">`);
  });

  it("moves a conditional format's range", async () => {
    const sheet = `${SHEET_OPEN}<sheetData/><conditionalFormatting sqref="B9:B11"><cfRule type="cellIs"/></conditionalFormatting></worksheet>`;

    expect(await shifted(sheet, INSERT)).toContain(`sqref="B9:B13"`);
  });

  it("moves every range of a conditional format that covers more than one", async () => {
    const sheet = `${SHEET_OPEN}<sheetData/><conditionalFormatting sqref="B12 D12:D14"/></worksheet>`;

    expect(await shifted(sheet, INSERT)).toContain(`sqref="B14 D14:D16"`);
  });

  it("moves a data validation, a hyperlink and a filter", async () => {
    const sheet = `${SHEET_OPEN}<sheetData/><autoFilter ref="A8:C12"/><dataValidation sqref="B12"/><hyperlink ref="A12"/></worksheet>`;

    const out = await shifted(sheet, INSERT);

    expect(out).toContain(`<autoFilter ref="A8:C14">`);
    expect(out).toContain(`<dataValidation sqref="B14">`);
    expect(out).toContain(`<hyperlink ref="A14">`);
  });

  it("moves the cell a frozen pane starts at", async () => {
    const sheet = `${SHEET_OPEN}<sheetViews><sheetView><pane ySplit="11" topLeftCell="A12" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>`;

    expect(await shifted(sheet, INSERT)).toContain(`topLeftCell="A14"`);
  });

  it("moves a page break", async () => {
    const sheet = `${SHEET_OPEN}<sheetData/><rowBreaks count="1"><brk id="12" max="16383" man="1"/></rowBreaks></worksheet>`;

    expect(await shifted(sheet, INSERT)).toContain(`<brk id="14"`);
  });

  it("leaves a row's column spans alone, since no column moved", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="12" spans="1:3"><c r="A12"><v>1</v></c></row></sheetData></worksheet>`;

    expect(await shifted(sheet, INSERT)).toContain(`<row r="14" spans="1:3">`);
  });

  // Anything in there can carry a reference in a form we do not read, and a stale
  // one is a sheet that looks right and is not.
  it("refuses a sheet holding an extension list", async () => {
    const sheet = `${SHEET_OPEN}<sheetData/><extLst><ext uri="{x}"/></extLst></worksheet>`;

    await expect(shifted(sheet, INSERT)).rejects.toThrow(
      "holds an extension list, whose contents cannot be moved with confidence",
    );
  });
});
