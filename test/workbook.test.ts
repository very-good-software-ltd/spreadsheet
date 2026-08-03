import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { Workbook } from "../src/workbook";
import type { Worksheet } from "../src/worksheet";
import { xlsx } from "./support/xlsx-fixture";

async function firstCellValue(worksheet: Worksheet): Promise<unknown> {
  for await (const row of worksheet.rows()) {
    return row.cells[0]?.value;
  }
  return undefined;
}

describe("Workbook", () => {
  it("lists worksheet names in document order", async () => {
    const workbook = await Workbook.open(
      xlsx([
        { name: "Summary", rows: [] },
        { name: "Data", rows: [] },
      ]),
    );

    expect(workbook.worksheetNames).toEqual(["Summary", "Data"]);
  });

  it("reports worksheet visibility", async () => {
    const workbook = await Workbook.open(
      xlsx([
        { name: "Visible", rows: [] },
        { name: "Hidden", rows: [], hidden: true },
      ]),
    );

    expect(workbook.worksheets).toEqual([
      { name: "Visible", hidden: false },
      { name: "Hidden", hidden: true },
    ]);
  });

  it("gets a worksheet by name, index, or the first one", async () => {
    const workbook = await Workbook.open(
      xlsx([
        { name: "First", rows: [["a"]] },
        { name: "Second", rows: [["b"]] },
      ]),
    );

    expect(await firstCellValue(workbook.worksheet("Second"))).toBe("b");
    expect(await firstCellValue(workbook.worksheet(1))).toBe("b");
    expect(await firstCellValue(workbook.firstWorksheet())).toBe("a");
  });

  it("throws for an out-of-range worksheet index", async () => {
    const workbook = await Workbook.open(xlsx([{ name: "Only", rows: [] }]));

    expect(() => workbook.worksheet(5)).toThrow(/not found/i);
  });

  it("rejects bytes that are not a zip at all", async () => {
    await expect(Workbook.open(new Uint8Array([1, 2, 3, 4, 5]))).rejects.toThrow(/not a valid xlsx/i);
  });

  it("rejects a zip that has no workbook part", async () => {
    const bytes = zipSync({ "hello.txt": strToU8("hi") });

    await expect(Workbook.open(bytes)).rejects.toThrow(/missing xl\/workbook\.xml/i);
  });

  it("names the worksheet when its data part is missing from the archive", async () => {
    const workbook = `<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const rels = `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`;
    const bytes = zipSync({
      "xl/workbook.xml": strToU8(workbook),
      "xl/_rels/workbook.xml.rels": strToU8(rels),
    });
    const opened = await Workbook.open(bytes);

    expect(() => opened.worksheet("Data").rows()).toThrow(/Data/);
  });

  it("treats an out-of-range shared string as empty rather than crashing", async () => {
    const workbook = `<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const rels = `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`;
    const sheet = `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>99</v></c></row></sheetData></worksheet>`;
    const bytes = zipSync({
      "xl/workbook.xml": strToU8(workbook),
      "xl/_rels/workbook.xml.rels": strToU8(rels),
      "xl/worksheets/sheet1.xml": strToU8(sheet),
      "xl/sharedStrings.xml": strToU8(`<sst><si><t>only one</t></si></sst>`),
    });
    const opened = await Workbook.open(bytes);

    const rows = [];
    for await (const row of opened.worksheet("Data").rows()) {
      rows.push(row);
    }

    expect(rows[0]?.cells).toEqual([{ ref: "A1", columnIndex: 0, type: "string", value: "" }]);
  });

  it("names the part when its XML is malformed", async () => {
    const bytes = zipSync({ "xl/workbook.xml": strToU8(`<workbook><sheets><sheet name="Data"`) });

    await expect(Workbook.open(bytes)).rejects.toThrow(/xl\/workbook\.xml/);
  });
});
