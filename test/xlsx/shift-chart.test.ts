import { describe, expect, it } from "vitest";
import { shiftChartReferences } from "../../src/xlsx/shift-chart";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { writeXmlEvent } from "../../src/xml/write-xml";

function streamOf(...chunks: readonly string[]): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      for (const chunk of chunks) {
        controller.enqueue(new TextEncoder().encode(chunk));
      }
      controller.close();
    },
  });
}

async function through(shift: RowShift, ...part: readonly string[]): Promise<string> {
  let out = "";
  for await (const batch of shiftChartReferences(new SaxesXmlReader().read(streamOf(...part)), shift)) {
    for (const event of batch) {
      out += writeXmlEvent(event);
    }
  }
  return out;
}

// A series as Excel writes one: its name from a cell, its categories from one
// column and its values from another, each with a cached copy of what it read.
function chartOver(name: string, categories: string, values: string): string {
  return `<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>${name}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Amount</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>${categories}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${values}</c:f><c:numCache><c:ptCount val="5"/></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
}

const referencesIn = (part: string): readonly string[] =>
  [...part.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((match) => match[1] ?? "");

// Three rows appear at row 6, which is where a region of rows 2 to 6 makes room.
const INSERT: RowShift = { sheet: "Data", at: 6, by: 3 };

// Rows 5 and 6 go away.
const REMOVE: RowShift = { sheet: "Data", at: 5, by: -2 };

describe("shiftChartReferences", () => {
  it("stretches a series range when rows appear inside it", async () => {
    const part = await through(INSERT, chartOver("Data!$C$1", "Data!$A$2:$A$6", "Data!$C$2:$C$6"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "Data!$A$2:$A$9", "Data!$C$2:$C$9"]);
  });

  it("closes a series range up when rows inside it go away", async () => {
    const part = await through(REMOVE, chartOver("Data!$C$1", "Data!$A$2:$A$6", "Data!$C$2:$C$6"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "Data!$A$2:$A$4", "Data!$C$2:$C$4"]);
  });

  it("moves a series reading wholly below the rows down", async () => {
    const part = await through(INSERT, chartOver("Data!$C$1", "Data!$A$20:$A$24", "Data!$C$20:$C$24"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "Data!$A$23:$A$27", "Data!$C$23:$C$27"]);
  });

  it("leaves a series reading another sheet alone", async () => {
    const part = await through(INSERT, chartOver("Other!$C$1", "Other!$A$2:$A$6", "Other!$C$2:$C$6"));

    expect(referencesIn(part)).toEqual(["Other!$C$1", "Other!$A$2:$A$6", "Other!$C$2:$C$6"]);
  });

  // A chart reference is a formula, so it can carry what Excel itself writes when
  // the rows a series read are deleted. A pivot's range attribute cannot, which is
  // why that one refuses and this one does not.
  it("writes a broken reference where the rows took the whole series with them", async () => {
    const part = await through(REMOVE, chartOver("Data!$C$1", "Data!$A$5:$A$6", "Data!$C$5:$C$6"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "#REF!", "#REF!"]);
  });

  // A whole column already covers every row, so rows arriving inside it change
  // nothing about what it reads.
  it("leaves a series reading a whole column alone", async () => {
    const part = await through(INSERT, chartOver("Data!$C$1", "Data!$A:$A", "Data!$C:$C"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "Data!$A:$A", "Data!$C:$C"]);
  });

  it("moves a reference arriving in pieces", async () => {
    const whole = chartOver("Data!$C$1", "Data!$A$2:$A$6", "Data!$C$2:$C$6");
    const split = whole.indexOf("Data!$C$2:$C$6") + 6;

    const part = await through(INSERT, whole.slice(0, split), whole.slice(split));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "Data!$A$2:$A$9", "Data!$C$2:$C$9"]);
  });

  it("leaves the rest of the chart as it was", async () => {
    const part = await through(INSERT, chartOver("Data!$C$1", "Data!$A$2:$A$6", "Data!$C$2:$C$6"));

    expect(part).toContain("<c:v>Amount</c:v>");
    expect(part).toContain('<c:ptCount val="5">');
  });

  it("refuses a series we cannot move with confidence", async () => {
    await expect(through(INSERT, chartOver("Data!$C$1", "Data!$A$2:$A$6", "Data:Other!$C$2"))).rejects.toThrow(
      "refers to a range of sheets",
    );
  });
});

// A chart built from a named range holds the name rather than the cells. The name
// moves in the workbook part, so moving it here as well would move it twice.
describe("shiftChartReferences, on a reference that is not cells", () => {
  it("leaves a series reading a defined name alone", async () => {
    const part = await through(INSERT, chartOver("Data!$C$1", "Data!Months", "Data!Amounts"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "Data!Months", "Data!Amounts"]);
  });

  it("leaves a series reading another workbook alone", async () => {
    const part = await through(INSERT, chartOver("Data!$C$1", "[1]Data!$A$2:$A$6", "[1]Data!$C$2:$C$6"));

    expect(referencesIn(part)).toEqual(["Data!$C$1", "[1]Data!$A$2:$A$6", "[1]Data!$C$2:$C$6"]);
  });

  it("moves a series naming a sheet whose name is quoted", async () => {
    const shift: RowShift = { sheet: "My Data", at: 6, by: 3 };
    const part = await through(shift, chartOver("'My Data'!$C$1", "'My Data'!$A$2:$A$6", "'My Data'!$C$2:$C$6"));

    expect(referencesIn(part)).toEqual(["'My Data'!$C$1", "'My Data'!$A$2:$A$9", "'My Data'!$C$2:$C$9"]);
  });
});
