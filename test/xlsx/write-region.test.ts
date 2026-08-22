import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/workbook";
import { xlsx } from "../support/xlsx-fixture";

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream as unknown as AsyncIterable<Uint8Array>) {
    chunks.push(chunk);
  }
  return new Uint8Array(chunks.flatMap((chunk) => [...chunk]));
}

async function cellsOf(bytes: Uint8Array): Promise<unknown[][]> {
  const workbook = await Workbook.open(bytes);
  const rows: unknown[][] = [];

  for await (const row of workbook.worksheet(0).rows()) {
    const values: unknown[] = [];
    for (const cell of row.cells) {
      values[cell.columnIndex] = cell.value;
    }
    rows[row.number - 1] = values;
  }

  return [...rows];
}

// A region three rows tall and three columns wide, over B4:D6, with the sheet
// holding a value in every cell it covers and in the cells around it.
function template() {
  const filled = ["x", "y", "z"];
  return xlsx(
    [
      {
        name: "Report",
        rows: [
          ["heading"],
          [],
          [],
          ["left", ...filled, "right"],
          ["left", ...filled, "right"],
          ["left", ...filled, "right"],
          ["total", "under"],
        ],
      },
    ],
    { definedNames: [{ name: "Data", target: "Report!$B$4:$D$6" }] },
  );
}

async function filledWith(write: (worksheet: ReturnType<Workbook["edit"]>) => void) {
  const workbook = await Workbook.open(template());
  const editor = workbook.edit();
  write(editor);
  return cellsOf(await bytesOf(editor.save()));
}

// The same region, with a total above it that reads the rows below it. That total
// cannot be written until the rows have been counted, so the writer counts them
// first and holds them, which is a different path through the region writer.
function templateWithTotalAbove() {
  return xlsx(
    [
      {
        name: "Report",
        rows: [
          [{ formula: "SUM(B4:B6)", cached: 0 }],
          [],
          [],
          ["left", "x", "y", "z", "right"],
          ["left", "x", "y", "z", "right"],
          ["left", "x", "y", "z", "right"],
        ],
      },
    ],
    { definedNames: [{ name: "Data", target: "Report!$B$4:$D$6" }] },
  );
}

async function filledUnderATotal(rows: readonly (readonly (string | number)[])[]) {
  const editor = (await Workbook.open(templateWithTotalAbove())).edit();
  editor.worksheet("Report").writeRegion("Data", rows);

  return cellsOf(await bytesOf(editor.save()));
}

describe("writing into a named region under a total that reads it", () => {
  it("gives each row its own values", async () => {
    const rows = await filledUnderATotal([
      ["North", "January", 1],
      ["South", "February", 2],
      ["East", "March", 3],
    ]);

    expect(rows.slice(3, 6)).toEqual([
      ["left", "North", "January", 1, "right"],
      ["left", "South", "February", 2, "right"],
      ["left", "East", "March", 3, "right"],
    ]);
  });

  it("gives each row its own values when the region grows", async () => {
    const rows = await filledUnderATotal([
      ["a", "b", 1],
      ["c", "d", 2],
      ["e", "f", 3],
      ["g", "h", 4],
    ]);

    expect(rows.slice(3, 7).map((row) => row[1])).toEqual(["a", "c", "e", "g"]);
  });
});

describe("writing into a named region", () => {
  it("writes rows from the region's first row and first column", async () => {
    const rows = await filledWith((editor) =>
      editor.worksheet("Report").writeRegion("Data", [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
      ]),
    );

    expect(rows[3]).toEqual(["left", 1, 2, 3, "right"]);
    expect(rows[5]).toEqual(["left", 7, 8, 9, "right"]);
  });

  it("leaves everything outside the region alone, and keeps what is below below", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", [[1, 2, 3]]));

    expect(rows[0]).toEqual(["heading"]);
    expect(rows[3]?.[0]).toBe("left");
    expect(rows[3]?.[4]).toBe("right");
    expect(rows[4]).toEqual(["total", "under"]);
  });

  it("takes away the rows of the region it was not given, rather than leaving them blank", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", [[1, 2, 3]]));

    expect(rows).toHaveLength(5);
  });

  it("makes room for more rows than the region held, pushing what is below down", async () => {
    const rows = await filledWith((editor) =>
      editor.worksheet("Report").writeRegion("Data", [
        [1, 1, 1],
        [2, 2, 2],
        [3, 3, 3],
        [4, 4, 4],
        [5, 5, 5],
      ]),
    );

    expect(rows[3]).toEqual(["left", 1, 1, 1, "right"]);
    expect(rows[6]).toEqual([undefined, 4, 4, 4]);
    expect(rows[8]).toEqual(["total", "under"]);
  });

  // Nothing is written with no data at all, but the region cannot vanish either: a
  // total written over it would have its whole range deleted and die.
  it("leaves one blank row when given nothing", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", []));

    expect(rows[3]).toEqual(["left", undefined, undefined, undefined, "right"]);
    expect(rows[4]).toEqual(["total", "under"]);
  });

  it("clears the columns a short row did not reach", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", [[1]]));

    expect(rows[3]).toEqual(["left", 1, undefined, undefined, "right"]);
  });

  // A gap the caller wrote deliberately still means leave it alone, where a row
  // that simply stops short is not a choice about the columns beyond it.
  it("leaves a cell the caller skipped with a gap", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", [[1, undefined, 3]]));

    expect(rows[3]).toEqual(["left", 1, "y", 3, "right"]);
  });

  it("refuses a row wider than the region", async () => {
    const workbook = await Workbook.open(template());
    const editor = workbook.edit();

    editor.worksheet("Report").writeRegion("Data", [[1, 2, 3, 4]]);

    await expect(bytesOf(editor.save())).rejects.toThrow('The name "Data" covers 3 columns and was given a row of 4');
  });

  it("finds a workbook-wide name from the workbook editor and writes into its own sheet", async () => {
    const workbook = await Workbook.open(template());
    const editor = workbook.edit();

    editor.writeRegion("Data", [[1, 2, 3]]);

    expect((await cellsOf(await bytesOf(editor.save())))[3]).toEqual(["left", 1, 2, 3, "right"]);
  });

  // The workbook part is one of the few we rewrite rather than copy, so the names
  // in it have to come through that rewrite intact or a second fill of the same
  // file would not find them.
  it("keeps the name in the written file, so the output can be filled again", async () => {
    const workbook = await Workbook.open(template());
    const editor = workbook.edit();
    editor.worksheet("Report").writeRegion("Data", [[1, 2, 3]]);

    const filled = await Workbook.open(await bytesOf(editor.save()));
    const again = filled.edit();
    again.worksheet("Report").writeRegion("Data", [[4, 5, 6]]);

    expect((await cellsOf(await bytesOf(again.save())))[3]).toEqual(["left", 4, 5, 6, "right"]);
  });

  it("moves the region's own name, so the output can be filled again with the same result", async () => {
    const workbook = await Workbook.open(template());
    const editor = workbook.edit();
    editor.worksheet("Report").writeRegion("Data", [
      [1, 1, 1],
      [2, 2, 2],
      [3, 3, 3],
      [4, 4, 4],
      [5, 5, 5],
    ]);

    const filled = await Workbook.open(await bytesOf(editor.save()));
    const again = filled.edit();
    again.worksheet("Report").writeRegion("Data", [[9, 9, 9]]);

    const rows = await cellsOf(await bytesOf(again.save()));
    expect(rows[3]).toEqual(["left", 9, 9, 9, "right"]);
    expect(rows[4]).toEqual(["total", "under"]);
  });

  it("refuses an unknown name at the call, not at save", async () => {
    const workbook = await Workbook.open(template());

    expect(() =>
      workbook
        .edit()
        .worksheet("Report")
        .writeRegion("Nope", [[1]]),
    ).toThrow('No name "Nope" on worksheet "Report" or in this workbook');
  });
});

// A table over B2:D6: its heading row, then four rows of data holding a previous
// run's figures, with the labels either side of it left out of the table.
function withTable() {
  const stale = [9, 9, 9];
  return xlsx([
    {
      name: "Report",
      rows: [
        [],
        ["", "Item", "Qty", "Amount", "keep me"],
        ["left", ...stale, "right"],
        ["left", ...stale, "right"],
        ["left", ...stale, "right"],
        ["left", ...stale, "right"],
      ],
      tables: [{ name: "Sales", ref: "B2:D6", headerRowCount: 1 }],
    },
  ]);
}

describe("writing into an Excel Table by name", () => {
  it("writes its data rows and leaves the heading row alone", async () => {
    const workbook = await Workbook.open(withTable());
    const editor = workbook.edit();

    editor.worksheet("Report").writeRegion("Sales", [
      [1, 2, 3],
      [4, 5, 6],
    ]);

    const rows = await cellsOf(await bytesOf(editor.save()));
    expect(rows[1]).toEqual(["", "Item", "Qty", "Amount", "keep me"]);
    expect(rows[2]).toEqual(["left", 1, 2, 3, "right"]);
    expect(rows[3]).toEqual(["left", 4, 5, 6, "right"]);
  });

  it("takes away the data rows it was not given", async () => {
    const workbook = await Workbook.open(withTable());
    const editor = workbook.edit();

    editor.worksheet("Report").writeRegion("Sales", [[1, 2, 3]]);

    expect(await cellsOf(await bytesOf(editor.save()))).toHaveLength(3);
  });

  it("finds the table from the workbook editor too", async () => {
    const workbook = await Workbook.open(withTable());
    const editor = workbook.edit();

    editor.writeRegion("Sales", [[1, 2, 3]]);

    expect((await cellsOf(await bytesOf(editor.save())))[2]).toEqual(["left", 1, 2, 3, "right"]);
  });
});

function tableExtentsIn(bytes: Uint8Array): (string | undefined)[] {
  const part = strFromU8(unzipSync(bytes)["xl/tables/table1.xml"] ?? new Uint8Array());
  return [...part.matchAll(/<(table|autoFilter)\b[^>]*\bref="([^"]+)"/g)].map((match) => match[2]);
}

async function grown(rows: readonly (readonly number[])[], totalsRowCount = 0) {
  const source = xlsx([
    {
      name: "Report",
      rows: [[], ["", "Item", "Qty", "Amount"], ["", 9, 9, 9], ["", 9, 9, 9]],
      tables: [{ name: "Sales", ref: totalsRowCount > 0 ? "B2:D5" : "B2:D4", headerRowCount: 1, totalsRowCount }],
    },
  ]);
  const editor = (await Workbook.open(source)).edit();
  editor.worksheet("Report").writeRegion("Sales", rows);

  return bytesOf(editor.save());
}

describe("growing an Excel Table to fit the rows it is given", () => {
  it("extends the table and its filter to cover the rows that did not fit", async () => {
    expect(
      tableExtentsIn(
        await grown([
          [1, 1, 1],
          [2, 2, 2],
          [3, 3, 3],
          [4, 4, 4],
        ]),
      ),
    ).toEqual(["B2:D6", "B2:D6"]);
  });

  it("puts those rows on the sheet past where the table used to end", async () => {
    const rows = await cellsOf(
      await grown([
        [1, 1, 1],
        [2, 2, 2],
        [3, 3, 3],
        [4, 4, 4],
      ]),
    );

    expect(rows[4]).toEqual([undefined, 3, 3, 3]);
    expect(rows[5]).toEqual(["", 4, 4, 4]);
  });

  it("pulls the table in when given fewer rows than it held", async () => {
    const bytes = await grown([[1, 1, 1]]);

    expect(tableExtentsIn(bytes)).toEqual(["B2:D3", "B2:D3"]);
    expect(await cellsOf(bytes)).toHaveLength(3);
  });

  it("leaves the table alone when the rows it was given fit", async () => {
    expect(
      tableExtentsIn(
        await grown([
          [1, 1, 1],
          [2, 2, 2],
        ]),
      ),
    ).toEqual(["B2:D4", "B2:D4"]);
  });

  // The totals row sits under the data, and now it simply moves down with
  // everything else rather than standing in the way.
  it("grows a table that has a totals row, pushing that row down", async () => {
    const bytes = await grown(
      [
        [1, 1, 1],
        [2, 2, 2],
        [3, 3, 3],
      ],
      1,
    );

    expect(tableExtentsIn(bytes)).toEqual(["B2:D6", "B2:D6"]);
  });
});
