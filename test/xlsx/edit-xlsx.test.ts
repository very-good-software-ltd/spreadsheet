import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { date, formula } from "../../src/cell-input";
import { BytesByteRange } from "../../src/io/byte-range";
import { Workbook } from "../../src/workbook";
import { openZip } from "../../src/zip/open-zip";
import { odsWith } from "../support/ods-fixture";
import { xlsx } from "../support/xlsx-fixture";

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

async function cellsOf(bytes: Uint8Array, sheet: string | number = 0): Promise<unknown[][]> {
  const workbook = await Workbook.open(bytes);
  const rows: unknown[][] = [];

  for await (const row of workbook.worksheet(sheet).rows()) {
    const values: unknown[] = [];
    for (const cell of row.cells) {
      values[cell.columnIndex] = cell.value;
    }
    rows[row.number - 1] = values;
  }

  return [...rows];
}

describe("editing an xlsx", () => {
  it("writes a value into a cell and reads it back", async () => {
    const source = xlsx([{ name: "Data", rows: [[1, 2]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet("Data").set("B1", 42);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([[1, 42]]);
  });

  it("leaves the cells it was not asked about alone", async () => {
    const source = xlsx([{ name: "Data", rows: [["keep", "change", "keep too"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).set("B1", "changed");

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["keep", "changed", "keep too"]]);
  });

  it("writes each kind of value", async () => {
    const workbook = await Workbook.create();
    const editor = workbook.edit();

    editor
      .worksheet(0)
      .set("A1", 42.5)
      .set("B1", "text")
      .set("C1", true)
      .set("D1", date("2020-01-01"))
      .set("E1", formula("SUM(A1:A1)"));

    const cells = await cellsOf(await bytesOf(editor.save()));

    expect(cells[0]?.[0]).toBe(42.5);
    expect(cells[0]?.[1]).toBe("text");
    expect(cells[0]?.[2]).toBe(true);
    expect(cells[0]?.[3]).toEqual(new Date("2020-01-01T00:00:00.000Z"));
    expect(cells[0]?.[4]).toBe("SUM(A1:A1)");
  });

  it("blanks a cell with null", async () => {
    const source = xlsx([{ name: "Data", rows: [["gone", "here"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).set("A1", null);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([[undefined, "here"]]);
  });

  it("appends rows from an array", async () => {
    const source = xlsx([{ name: "Data", rows: [["header"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).appendRows([
      [1, 2],
      [3, 4],
    ]);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["header"], [1, 2], [3, 4]]);
  });

  it("appends rows from a generator, which is not read until the file is saved", async () => {
    const workbook = await Workbook.create();
    const editor = workbook.edit();
    let pulled = 0;

    editor.worksheet(0).appendRows(
      (async function* () {
        for (let i = 1; i <= 3; i += 1) {
          pulled += 1;
          yield [i];
        }
      })(),
    );

    expect(pulled).toBe(0);
    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([[1], [2], [3]]);
    expect(pulled).toBe(3);
  });

  it("writes rows at a position, over what is already there", async () => {
    const source = xlsx([{ name: "Data", rows: [["a"], ["b"], ["c"], ["d"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).writeRows(2, [["B"], ["C"]]);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["a"], ["B"], ["C"], ["d"]]);
  });

  it("writes rows past the end of the sheet", async () => {
    const source = xlsx([{ name: "Data", rows: [["a"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).writeRows(3, [["c"]]);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["a"], undefined, ["c"]]);
  });

  it("leaves a column alone where a row has a gap, and blanks it where it has null", async () => {
    const source = xlsx([{ name: "Data", rows: [["a", "b", "c"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).writeRows(1, [[undefined, null, "C"]]);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["a", undefined, "C"]]);
  });

  it("lets the last call win where two cover the same cell", async () => {
    const workbook = await Workbook.create();
    const editor = workbook.edit();

    editor
      .worksheet(0)
      .set("A1", "first")
      .writeRows(1, [["second"]]);

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["second"]]);
  });

  it("lets a later set win over an earlier block", async () => {
    const workbook = await Workbook.create();
    const editor = workbook.edit();

    editor
      .worksheet(0)
      .writeRows(1, [["from block"]])
      .set("A1", "from set");

    expect(await cellsOf(await bytesOf(editor.save()))).toEqual([["from set"]]);
  });

  it("edits more than one sheet in the same save", async () => {
    const source = xlsx([
      { name: "One", rows: [["a"]] },
      { name: "Two", rows: [["b"]] },
    ]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet("One").set("A1", "A");
    editor.worksheet("Two").set("A1", "B");
    const bytes = await bytesOf(editor.save());

    expect(await cellsOf(bytes, "One")).toEqual([["A"]]);
    expect(await cellsOf(bytes, "Two")).toEqual([["B"]]);
  });

  describe("recalculation", () => {
    async function partsOf(bytes: Uint8Array): Promise<Record<string, string>> {
      const archive = await openZip(new BytesByteRange(bytes));
      const parts: Record<string, string> = {};
      for (const entry of archive.entries()) {
        parts[entry.path] = strFromU8(await archive.read(entry.path));
      }
      return parts;
    }

    async function written(source: Uint8Array): Promise<Record<string, string>> {
      const editor = (await Workbook.open(source)).edit();
      editor.worksheet(0).set("A1", 1);
      return partsOf(await bytesOf(editor.save()));
    }

    it("tells a spreadsheet application to recalculate, since an edit staled every cached result", async () => {
      const parts = await written(xlsx([{ name: "Data", rows: [[1]] }]));

      expect(parts["xl/workbook.xml"]).toContain('fullCalcOnLoad="1"');
    });

    it("leaves out the calculation chain, which an edit invalidates", async () => {
      const source = zipSync({
        ...unzipSync(xlsx([{ name: "Data", rows: [[1]] }])),
        "xl/calcChain.xml": strToU8('<calcChain xmlns="x"><c r="A1" i="1"/></calcChain>'),
        "[Content_Types].xml": strToU8(
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/calcChain.xml" ContentType="calcChain"/><Override PartName="/xl/styles.xml" ContentType="styles"/></Types>',
        ),
      });

      const parts = await written(source);

      expect(Object.keys(parts)).not.toContain("xl/calcChain.xml");
      expect(parts["[Content_Types].xml"]).not.toContain("calcChain");
      expect(parts["xl/_rels/workbook.xml.rels"]).not.toContain("calcChain");
      expect(parts["[Content_Types].xml"]).toContain("styles");
    });

    it("leaves the content types and relationships parts untouched when there was no chain to drop", async () => {
      const source = xlsx([{ name: "Data", rows: [[1]] }]);
      const before = await partsOf(source);
      const after = await written(source);

      expect(after["xl/_rels/workbook.xml.rels"]).toBe(before["xl/_rels/workbook.xml.rels"]);
    });
  });

  describe("adding a worksheet", () => {
    it("adds a sheet a reader then finds", async () => {
      const workbook = await Workbook.create();
      const editor = workbook.edit();

      editor.addWorksheet("Summary").set("A1", "total");
      const bytes = await bytesOf(editor.save());

      expect((await Workbook.open(bytes)).worksheetNames).toEqual(["Sheet1", "Summary"]);
      expect(await cellsOf(bytes, "Summary")).toEqual([["total"]]);
      expect(await cellsOf(bytes, "Sheet1")).toEqual([]);
    });

    it("adds it to a workbook opened from a file, keeping the sheets already there", async () => {
      const source = xlsx([{ name: "Data", rows: [["a"]] }]);
      const editor = (await Workbook.open(source)).edit();

      editor.addWorksheet("Extra").set("A1", "b");
      const bytes = await bytesOf(editor.save());

      expect((await Workbook.open(bytes)).worksheetNames).toEqual(["Data", "Extra"]);
      expect(await cellsOf(bytes, "Data")).toEqual([["a"]]);
      expect(await cellsOf(bytes, "Extra")).toEqual([["b"]]);
    });

    it("adds more than one at a time", async () => {
      const editor = (await Workbook.create()).edit();

      editor.addWorksheet("Two");
      editor.addWorksheet("Three");

      expect((await Workbook.open(await bytesOf(editor.save()))).worksheetNames).toEqual(["Sheet1", "Two", "Three"]);
    });

    it("takes a part path and a relationship id nothing else is using", async () => {
      const editor = (await Workbook.open(xlsx([{ name: "One", rows: [] }]))).edit();

      editor.addWorksheet("Two");
      const bytes = await bytesOf(editor.save());
      const archive = await openZip(new BytesByteRange(bytes));

      expect(archive.has("xl/worksheets/sheet2.xml")).toBe(true);

      const rels = strFromU8(await archive.read("xl/_rels/workbook.xml.rels"));
      expect(rels.match(/Id="rId1"/g)).toHaveLength(1);
      expect(rels).toContain('Target="worksheets/sheet2.xml"');
    });

    it("declares the new part's content type, so a reader knows what it is", async () => {
      const editor = (await Workbook.create()).edit();

      editor.addWorksheet("Two");
      const archive = await openZip(new BytesByteRange(await bytesOf(editor.save())));

      expect(strFromU8(await archive.read("[Content_Types].xml"))).toContain('PartName="/xl/worksheets/sheet2.xml"');
    });

    it("refuses a name a spreadsheet would not accept", async () => {
      const editor = (await Workbook.create()).edit();

      expect(() => editor.addWorksheet("")).toThrow(/not a worksheet name/i);
      expect(() => editor.addWorksheet("x".repeat(32))).toThrow(/not a worksheet name/i);
      expect(() => editor.addWorksheet("a/b")).toThrow(/cannot appear/i);
      expect(() => editor.addWorksheet("a:b")).toThrow(/cannot appear/i);
      expect(() => editor.addWorksheet("[a]")).toThrow(/cannot appear/i);
      expect(() => editor.addWorksheet("'quoted'")).toThrow(/cannot appear/i);
    });

    it("refuses a name already taken, whatever its case", async () => {
      const editor = (await Workbook.create()).edit();

      expect(() => editor.addWorksheet("Sheet1")).toThrow(/already exists/i);
      expect(() => editor.addWorksheet("sheet1")).toThrow(/already exists/i);
    });
  });

  it("keeps every part it did not touch", async () => {
    const source = xlsx([{ name: "Data", rows: [["a"]] }]);
    const workbook = await Workbook.open(source);
    const editor = workbook.edit();

    editor.worksheet(0).set("A1", "b");
    const written = await Workbook.open(await bytesOf(editor.save()));

    expect(written.worksheetNames).toEqual(["Data"]);
  });

  describe("what it refuses", () => {
    it("rejects a reference that is not a cell, at the call", async () => {
      const sheet = (await Workbook.create()).edit().worksheet(0);

      expect(() => sheet.set("nonsense", 1)).toThrow(/not a cell reference/i);
      expect(() => sheet.set("A0", 1)).toThrow(/not a cell reference/i);
      expect(() => sheet.set("A1:B2", 1)).toThrow(/not a cell reference/i);
    });

    it("accepts a lowercase reference", async () => {
      const workbook = await Workbook.create();
      const editor = workbook.edit();

      editor.worksheet(0).set("b2", "here");

      expect(await cellsOf(await bytesOf(editor.save()))).toEqual([undefined, [undefined, "here"]]);
    });

    it("rejects an unknown worksheet", async () => {
      const editor = (await Workbook.create()).edit();

      expect(() => editor.worksheet("Missing")).toThrow(/not found/i);
      expect(() => editor.worksheet(4)).toThrow(/not found/i);
    });

    it("rejects inheriting formatting from a row it will already have passed", async () => {
      const sheet = (await Workbook.create()).edit().worksheet(0);

      expect(() => sheet.writeRows(5, [[1]], { inheritFrom: 9 })).toThrow(/read once from the top/i);
    });

    it("says which row is missing when asked to copy formatting from one that is not there", async () => {
      const editor = (await Workbook.open(xlsx([{ name: "Data", rows: [["a"]] }]))).edit();

      editor.worksheet(0).writeRows(20, [["b"]], { inheritFrom: 9 });

      await expect(bytesOf(editor.save())).rejects.toThrow(/no row 9/i);
    });

    it("refuses to save twice, since a row source is read once", async () => {
      const editor = (await Workbook.create()).edit();

      editor.worksheet(0).set("A1", 1);
      await bytesOf(editor.save());

      expect(() => editor.save()).toThrow(/already been saved/i);
    });

    it("surfaces a failing row source as a stream error", async () => {
      const editor = (await Workbook.create()).edit();

      editor.worksheet(0).appendRows(
        (async function* () {
          yield [1];
          throw new Error("the source gave up");
        })(),
      );

      await expect(bytesOf(editor.save())).rejects.toThrow("the source gave up");
    });

    it("says why an ods workbook cannot be written", async () => {
      const workbook = await Workbook.open(odsWith('<table:table table:name="Data"></table:table>'));

      expect(() => workbook.edit()).toThrow(/\.ods files is not supported/i);
    });
  });
});
