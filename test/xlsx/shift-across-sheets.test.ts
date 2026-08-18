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

async function formulaOn(bytes: Uint8Array, sheet: string): Promise<unknown> {
  const workbook = await Workbook.open(bytes);
  for await (const row of workbook.worksheet(sheet).rows()) {
    for (const cell of row.cells) {
      if (cell.type === "formula") {
        return cell.value;
      }
    }
  }
  return undefined;
}

// Data on Report over rows 3 to 5, and a Summary sheet whose total reads them.
function template(summaryFormula: string) {
  return xlsx(
    [
      { name: "Report", rows: [[], [], [1], [2], [3]] },
      { name: "Summary", rows: [[{ formula: summaryFormula, cached: 0 }]] },
    ],
    { definedNames: [{ name: "Data", target: "Report!$A$3:$A$5" }] },
  );
}

async function afterWriting(rows: readonly (readonly number[])[], summaryFormula: string) {
  const editor = (await Workbook.open(template(summaryFormula))).edit();
  editor.worksheet("Report").writeRegion("Data", rows);

  return bytesOf(editor.save());
}

describe("moving rows on one worksheet", () => {
  it("moves a formula on another worksheet that reads them", async () => {
    const bytes = await afterWriting([[1], [2], [3], [4], [5]], "SUM(Report!A3:A5)");

    expect(await formulaOn(bytes, "Summary")).toBe("SUM(Report!A3:A7)");
  });

  it("shrinks one when the rows go away", async () => {
    const bytes = await afterWriting([[1]], "SUM(Report!A3:A5)");

    expect(await formulaOn(bytes, "Summary")).toBe("SUM(Report!A3:A3)");
  });

  it("leaves a formula naming its own sheet alone", async () => {
    const bytes = await afterWriting([[1], [2], [3], [4], [5]], "SUM(A3:A5)");

    expect(await formulaOn(bytes, "Summary")).toBe("SUM(A3:A5)");
  });

  it("leaves a formula naming a third sheet alone", async () => {
    const bytes = await afterWriting([[1], [2], [3], [4], [5]], "SUM(Elsewhere!A3:A5)");

    expect(await formulaOn(bytes, "Summary")).toBe("SUM(Elsewhere!A3:A5)");
  });
});

describe("moving rows on two worksheets in one save", () => {
  // Each sheet's total reads the other's rows, so neither move can be applied
  // without the other being known first.
  function pair() {
    return xlsx(
      [
        { name: "First", rows: [[{ formula: "SUM(Second!A3:A5)", cached: 0 }], [], [1], [2], [3]] },
        { name: "Second", rows: [[{ formula: "SUM(First!A3:A5)", cached: 0 }], [], [1], [2], [3]] },
      ],
      {
        definedNames: [
          { name: "FirstData", target: "First!$A$3:$A$5" },
          { name: "SecondData", target: "Second!$A$3:$A$5" },
        ],
      },
    );
  }

  it("moves each sheet's formulas for the other sheet's move", async () => {
    const editor = (await Workbook.open(pair())).edit();
    editor.worksheet("First").writeRegion("FirstData", [[1], [2], [3], [4], [5]]);
    editor.worksheet("Second").writeRegion("SecondData", [[1]]);

    const bytes = await bytesOf(editor.save());

    expect(await formulaOn(bytes, "First")).toBe("SUM(Second!A3:A3)");
    expect(await formulaOn(bytes, "Second")).toBe("SUM(First!A3:A7)");
  });
});
