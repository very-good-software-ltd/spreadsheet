import { describe, expect, it } from "vitest";
import type { CellInput } from "../../src/cell-input";
import { formula } from "../../src/cell-input";
import type { DateStyles } from "../../src/xlsx/write-sheet";
import { type RowCells, type RowEdit, type SheetWritePlan, writeSheetPart } from "../../src/xlsx/write-sheet";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { XML_DECLARATION } from "../../src/xml/write-xml";

const SHEET_OPEN = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;

// Every date lands on the same style here, so a test can tell that a date was
// styled without depending on how styles.xml is rewritten.
const dateStyles: DateStyles = { forDate: () => "99" };

function streamOf(text: string): ReadableStream<Uint8Array> {
  const bytes = new TextEncoder().encode(text);
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(bytes);
      controller.close();
    },
  });
}

function cells(values: readonly CellInput[], firstColumn = 0): RowCells {
  return new Map(values.map((value, index) => [firstColumn + index, value]));
}

function plan(positioned: readonly RowEdit[] = [], appended: readonly RowCells[] = []): SheetWritePlan {
  return {
    positioned: (async function* () {
      yield* positioned;
    })(),
    appended: (async function* () {
      yield* appended;
    })(),
  };
}

async function rewrite(sheet: string, sheetPlan: SheetWritePlan = plan()): Promise<string> {
  const events = new SaxesXmlReader().read(streamOf(sheet));
  let out = "";
  for await (const chunk of writeSheetPart(events, sheetPlan, { dateStyles, date1904: false })) {
    out += chunk;
  }
  return out.slice(XML_DECLARATION.length);
}

describe("writeSheetPart", () => {
  it("passes a sheet it has no edits for straight through", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="3" t="s"><v>0</v></c></row></sheetData></worksheet>`;

    expect(await rewrite(sheet)).toBe(sheet);
  });

  it("keeps everything after the rows, which is where a template's formatting lives", async () => {
    const tail = `<mergeCells count="1"><mergeCell ref="A1:B1"></mergeCell></mergeCells><conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="0"></cfRule></conditionalFormatting>`;
    const sheet = `${SHEET_OPEN}<cols><col min="1" max="1" width="30"></col></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${tail}</worksheet>`;

    expect(await rewrite(sheet, plan([{ number: 1, cells: cells([5]) }]))).toContain(tail);
  });

  it("drops the dimension, whose extent a single pass cannot know before the rows", async () => {
    const sheet = `${SHEET_OPEN}<dimension ref="A1:B2"></dimension><sheetData></sheetData></worksheet>`;

    expect(await rewrite(sheet)).toBe(`${SHEET_OPEN}<sheetData></sheetData></worksheet>`);
  });

  it("replaces a value and keeps the cell's style", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="7"><v>1</v></c></row></sheetData></worksheet>`;

    expect(await rewrite(sheet, plan([{ number: 1, cells: cells([42]) }]))).toBe(
      `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="7"><v>42</v></c></row></sheetData></worksheet>`,
    );
  });

  it("leaves the other cells in an edited row alone", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2" t="s"><v>4</v></c><c r="C1" s="3"><v>3</v></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([{ number: 1, cells: new Map([[1, 9]]) }]));

    expect(result).toContain(`<c r="A1" s="1"><v>1</v></c>`);
    expect(result).toContain(`<c r="B1" s="2"><v>9</v></c>`);
    expect(result).toContain(`<c r="C1" s="3"><v>3</v></c>`);
  });

  it("drops the formula when a plain value overwrites a formula cell", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="4"><f>SUM(B1:C1)</f><v>7</v></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([{ number: 1, cells: cells([42]) }]));

    expect(result).not.toContain("<f>");
    expect(result).toContain(`<c r="A1" s="4"><v>42</v></c>`);
  });

  it("inserts a new cell in column order, not at the end of the row", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row></sheetData></worksheet>`;

    expect(await rewrite(sheet, plan([{ number: 1, cells: new Map([[1, 2]]) }]))).toBe(
      `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData></worksheet>`,
    );
  });

  it("writes a row that the sheet does not have, in row order", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([{ number: 2, cells: cells([2]) }]));

    expect(result).toBe(
      `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>`,
    );
  });

  it("writes rows past the last existing row", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([{ number: 5, cells: cells([5]) }]));

    expect(result).toContain(`<row r="5"><c r="A5"><v>5</v></c></row></sheetData>`);
  });

  it("appends rows after the last existing row", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="4"><c r="A4"><v>4</v></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([], [cells([5, 6]), cells([7, 8])]));

    expect(result).toContain(
      `<row r="5"><c r="A5"><v>5</v></c><c r="B5"><v>6</v></c></row><row r="6"><c r="A6"><v>7</v></c><c r="B6"><v>8</v></c></row></sheetData>`,
    );
  });

  it("appends from row one when the sheet has no rows", async () => {
    const sheet = `${SHEET_OPEN}<sheetData></sheetData></worksheet>`;

    expect(await rewrite(sheet, plan([], [cells([1])]))).toBe(
      `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
    );
  });

  it("blanks a cell with null and keeps its formatting", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="6" t="s"><v>2</v></c></row></sheetData></worksheet>`;

    expect(await rewrite(sheet, plan([{ number: 1, cells: cells([null]) }]))).toBe(
      `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="6"></c></row></sheetData></worksheet>`,
    );
  });

  it("copies cell formatting from a nominated row onto rows written past it", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="6" ht="24" customHeight="1"><c r="A6" s="11"></c><c r="B6" s="12"></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([{ number: 7, cells: cells([1, 2]), inheritFrom: 6 }]));

    expect(result).toContain(
      `<row r="7" ht="24" customHeight="1"><c r="A7" s="11"><v>1</v></c><c r="B7" s="12"><v>2</v></c></row>`,
    );
  });

  it("prefers a cell's own style over an inherited one", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1" s="11"></c></row><row r="2"><c r="A2" s="22"></c></row></sheetData></worksheet>`;

    const result = await rewrite(sheet, plan([{ number: 2, cells: cells([5]), inheritFrom: 1 }]));

    expect(result).toContain(`<c r="A2" s="22"><v>5</v></c>`);
  });

  it("recomputes spans when it changes a row's extent, and leaves it alone otherwise", async () => {
    const sheet = `${SHEET_OPEN}<sheetData><row r="1" spans="1:1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;

    expect(await rewrite(sheet, plan([{ number: 1, cells: new Map([[2, 3]]) }]))).toContain(`<row r="1" spans="1:3">`);
    expect(await rewrite(sheet)).toContain(`<row r="1" spans="1:1">`);
  });

  describe("cell values", () => {
    async function written(value: CellInput): Promise<string> {
      const sheet = `${SHEET_OPEN}<sheetData><row r="1"><c r="A1"><v>0</v></c></row></sheetData></worksheet>`;
      const result = await rewrite(sheet, plan([{ number: 1, cells: cells([value]) }]));
      return result.slice(result.indexOf("<c "), result.indexOf("</row>"));
    }

    it("writes a number with no type, which is what the absence of one means", async () => {
      expect(await written(42.5)).toBe(`<c r="A1"><v>42.5</v></c>`);
    });

    it("writes a string inline, so the shared strings table is never touched", async () => {
      expect(await written("hello")).toBe(`<c r="A1" t="inlineStr"><is><t>hello</t></is></c>`);
    });

    it("marks a string whose whitespace matters, which a reader would otherwise trim", async () => {
      expect(await written("  padded  ")).toBe(
        `<c r="A1" t="inlineStr"><is><t xml:space="preserve">  padded  </t></is></c>`,
      );
    });

    it("escapes a string's markup characters", async () => {
      expect(await written("a & b < c")).toBe(`<c r="A1" t="inlineStr"><is><t>a &amp; b &lt; c</t></is></c>`);
    });

    it("writes a boolean as one or zero", async () => {
      expect(await written(true)).toBe(`<c r="A1" t="b"><v>1</v></c>`);
      expect(await written(false)).toBe(`<c r="A1" t="b"><v>0</v></c>`);
    });

    it("writes a date as its serial with a date style", async () => {
      expect(await written(new Date("2020-01-01T00:00:00.000Z"))).toBe(`<c r="A1" s="99"><v>43831</v></c>`);
    });

    it("writes a formula with no cached result, leaving the value to be recalculated", async () => {
      expect(await written(formula("SUM(B1:C1)"))).toBe(`<c r="A1"><f>SUM(B1:C1)</f></c>`);
    });

    it("refuses a string holding a character XML cannot represent", async () => {
      await expect(written("bad\u0000value")).rejects.toThrow(/cannot be written/i);
      await expect(written("tab\tnewline\n")).resolves.toContain("tab");
    });

    it("refuses a number a spreadsheet has no way to hold", async () => {
      await expect(written(Number.NaN)).rejects.toThrow(/NaN/i);
      await expect(written(Number.POSITIVE_INFINITY)).rejects.toThrow(/finite/i);
    });
  });
});
