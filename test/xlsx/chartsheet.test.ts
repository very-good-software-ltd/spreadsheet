import { readFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { BytesByteRange } from "../../src/io/byte-range";
import { Workbook } from "../../src/workbook";
import { readWorkbook } from "../../src/xlsx/read-workbook";
import { createXmlReader } from "../../src/xml/create-xml-reader";
import { openZip } from "../../src/zip/open-zip";
import { type SheetInput, xlsx } from "../support/xlsx-fixture";

const CHART_SHEET_BETWEEN: readonly SheetInput[] = [
  { name: "Data", rows: [[1]] },
  { name: "Chart1", rows: [], chartsheet: true },
  { name: "Dashboard", rows: [[2]] },
];

const CHART_SHEET_LAST: readonly SheetInput[] = [
  { name: "Data", rows: [[1]] },
  { name: "Dashboard", rows: [[2]] },
  { name: "Chart1", rows: [], chartsheet: true },
];

async function workbookInfo(bytes: Uint8Array) {
  return readWorkbook(await openZip(new BytesByteRange(bytes)), createXmlReader());
}

describe("a chart sheet", () => {
  it("is not one of the worksheets", async () => {
    const info = await workbookInfo(xlsx(CHART_SHEET_BETWEEN));

    expect(info.worksheets.map((sheet) => sheet.name)).toEqual(["Data", "Dashboard"]);
  });

  // A name's scope is a position in the workbook's own sheet order, which counts
  // every sheet, so a chart sheet left out of the worksheets still has to be
  // counted when the position is resolved.
  it("still counts when the sheet a name is scoped to is resolved", async () => {
    const info = await workbookInfo(
      xlsx(CHART_SHEET_BETWEEN, { definedNames: [{ name: "Total", target: "Dashboard!$A$1", scope: 2 }] }),
    );

    expect(info.definedNames[0]?.scope).toBe("Dashboard");
  });

  it("still counts towards the highest sheet id", async () => {
    const info = await workbookInfo(xlsx(CHART_SHEET_LAST));

    expect(info.highestSheetId).toBe(3);
  });

  it("is not listed by a workbook", async () => {
    const workbook = await Workbook.open(xlsx(CHART_SHEET_BETWEEN));

    expect(workbook.worksheetNames).toEqual(["Data", "Dashboard"]);
  });

  it("cannot be asked for its rows", async () => {
    const workbook = await Workbook.open(xlsx(CHART_SHEET_BETWEEN));

    expect(() => workbook.worksheet("Chart1")).toThrow("Worksheet not found: Chart1");
  });

  it("leaves the worksheets after it addressable by position", async () => {
    const workbook = await Workbook.open(xlsx(CHART_SHEET_BETWEEN));
    const rows = [];

    for await (const row of workbook.worksheet(1).rows()) {
      rows.push(row);
    }

    expect(rows[0]?.cells[0]?.value).toBe(2);
  });

  it("cannot be edited", async () => {
    const workbook = await Workbook.open(xlsx(CHART_SHEET_BETWEEN));

    expect(() => workbook.edit().worksheet("Chart1")).toThrow("Worksheet not found: Chart1");
  });

  // The id has to avoid every sheet's, and a chart sheet holds one like any
  // other, so an added worksheet cannot take the number it already has.
  it("keeps its sheet id from being handed to an added worksheet", async () => {
    const workbook = await Workbook.open(xlsx(CHART_SHEET_LAST));
    const editor = workbook.edit();

    editor.addWorksheet("Added").appendRows([[1]]);

    const saved = await bytesOf(await editor.save());
    const reopened = await workbookInfo(saved);

    expect(reopened.highestSheetId).toBe(4);
  });

  // Dropping a sheet is a positive statement about what it is, and a relationship
  // naming no type makes no such statement.
  it("does not take a sheet whose relationship names no type", async () => {
    const workbook = `<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const rels = `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`;
    const info = await workbookInfo(
      zipSync({ "xl/workbook.xml": strToU8(workbook), "xl/_rels/workbook.xml.rels": strToU8(rels) }),
    );

    expect(info.worksheets.map((sheet) => sheet.name)).toEqual(["Data"]);
  });

  it("is left out of a file Excel wrote", async () => {
    const workbook = await Workbook.open(readFileSync("test/fixtures/chart-template.xlsx"));

    expect(workbook.worksheetNames).toEqual(["Data", "Dashboard"]);
  });
});

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;

  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    chunks.push(value);
    total += value.length;
  }

  const all = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    all.set(chunk, offset);
    offset += chunk.length;
  }

  return all;
}
