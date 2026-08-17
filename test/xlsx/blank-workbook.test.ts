import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/workbook";
import { blankXlsxArchive } from "../../src/xlsx/blank-workbook";
import { readXlsx } from "../../src/xlsx/read-xlsx";

async function collect<T>(source: AsyncIterable<T>): Promise<T[]> {
  const collected: T[] = [];
  for await (const item of source) {
    collected.push(item);
  }
  return collected;
}

describe("blankXlsxArchive", () => {
  it("holds every part a reader needs to make sense of it", () => {
    const archive = blankXlsxArchive();

    expect(
      archive
        .entries()
        .map((entry) => entry.path)
        .sort(),
    ).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/_rels/workbook.xml.rels",
      "xl/styles.xml",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
    ]);
  });

  it("stores its parts uncompressed, so copying one through needs no inflating", () => {
    const entry = blankXlsxArchive().storedEntry("xl/workbook.xml");

    expect(entry.method).toBe(0);
    expect(entry.compressedSize).toBe(entry.uncompressedSize);
    expect(entry.crc32).toBeGreaterThan(0);
  });

  it("reads as a workbook with one empty sheet", async () => {
    const data = await readXlsx(blankXlsxArchive());

    expect(data.worksheets).toEqual([{ name: "Sheet1", hidden: false }]);
    expect(await collect(data.openRows(0))).toEqual([]);
  });

  it("declares a styles part, so writing a date never has to create one", () => {
    expect(blankXlsxArchive().has("xl/styles.xml")).toBe(true);
  });
});

describe("Workbook.create", () => {
  it("hands back a workbook with one sheet, the same shape as an opened one", async () => {
    const workbook = await Workbook.create();

    expect(workbook.worksheetNames).toEqual(["Sheet1"]);
    expect(await collect(workbook.firstWorksheet().rows())).toEqual([]);
  });
});
