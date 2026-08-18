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

  it("leaves everything outside the region alone", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", [[1, 2, 3]]));

    expect(rows[0]).toEqual(["heading"]);
    expect(rows[3]?.[0]).toBe("left");
    expect(rows[3]?.[4]).toBe("right");
    expect(rows[6]).toEqual(["total", "under"]);
  });

  it("clears the rows of the region it was not given, so nothing stale is left behind", async () => {
    const rows = await filledWith((editor) => editor.worksheet("Report").writeRegion("Data", [[1, 2, 3]]));

    expect(rows[4]).toEqual(["left", undefined, undefined, undefined, "right"]);
    expect(rows[5]).toEqual(["left", undefined, undefined, undefined, "right"]);
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

  it("refuses more rows than the region holds rather than writing past it", async () => {
    const workbook = await Workbook.open(template());
    const editor = workbook.edit();

    editor.worksheet("Report").writeRegion("Data", [[1], [2], [3], [4]]);

    await expect(bytesOf(editor.save())).rejects.toThrow('The name "Data" covers 3 rows and was given more');
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
