import { describe, expect, it } from "vitest";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { shiftPivotLocation, shiftPivotSource, withCacheRefreshedOnLoad } from "../../src/xlsx/shift-pivot";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { writeXmlEvent } from "../../src/xml/write-xml";
import type { XmlEvent } from "../../src/xml/xml-reader";

type Shifter = (events: AsyncIterable<readonly XmlEvent[]>, shift: RowShift) => AsyncIterable<readonly XmlEvent[]>;

function streamOf(text: string): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(new TextEncoder().encode(text));
      controller.close();
    },
  });
}

async function through(part: string, apply: Shifter, shift: RowShift): Promise<string> {
  let out = "";
  for await (const batch of apply(new SaxesXmlReader().read(streamOf(part)), shift)) {
    for (const event of batch) {
      out += writeXmlEvent(event);
    }
  }
  return out;
}

function cacheOver(extent: string): string {
  return `<pivotCacheDefinition xmlns="m" recordCount="5"><cacheSource type="worksheet"><worksheetSource ref="${extent}" sheet="Data"/></cacheSource></pivotCacheDefinition>`;
}

function pivotAt(extent: string): string {
  return `<pivotTableDefinition xmlns="m" name="PivotTable1"><location ref="${extent}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>`;
}

const refOf = (part: string): string | undefined => /\bref="([^"]+)"/.exec(part)?.[1];

// Two rows appear at row 11.
const INSERT: RowShift = { sheet: "Data", at: 11, by: 2 };

// Rows 10 and 11 go away.
const REMOVE: RowShift = { sheet: "Data", at: 10, by: -2 };

describe("shiftPivotSource", () => {
  it("stretches the source range when rows appear inside it", async () => {
    expect(refOf(await through(cacheOver("A1:C15"), shiftPivotSource, INSERT))).toBe("A1:C17");
  });

  it("closes the source range up when rows inside it go away", async () => {
    expect(refOf(await through(cacheOver("A1:C15"), shiftPivotSource, REMOVE))).toBe("A1:C13");
  });

  it("leaves a source range above the rows alone", async () => {
    expect(refOf(await through(cacheOver("A1:C5"), shiftPivotSource, INSERT))).toBe("A1:C5");
  });

  it("moves a source range wholly below the rows down", async () => {
    expect(refOf(await through(cacheOver("A12:C15"), shiftPivotSource, INSERT))).toBe("A14:C17");
  });

  // A formula would read #REF! here, which an attribute holding a range has no
  // spelling for, so this is where a save gives up rather than writing nonsense.
  it("refuses when the rows take the whole source range with them", async () => {
    await expect(through(cacheOver("A10:C11"), shiftPivotSource, REMOVE)).rejects.toThrow(
      'The rows being taken out of "Data" take all of the range a pivot table reads with them',
    );
  });

  it("leaves the rest of the definition as it was", async () => {
    expect(await through(cacheOver("A1:C15"), shiftPivotSource, INSERT)).toContain('recordCount="5"');
  });
});

describe("shiftPivotLocation", () => {
  it("moves a pivot drawn below the rows down", async () => {
    expect(refOf(await through(pivotAt("A12:B16"), shiftPivotLocation, INSERT))).toBe("A14:B18");
  });

  it("leaves a pivot drawn above the rows alone", async () => {
    expect(refOf(await through(pivotAt("A3:B7"), shiftPivotLocation, INSERT))).toBe("A3:B7");
  });

  // These count from the top left of the ref rather than from the sheet, so moving
  // the block leaves them alone.
  it("leaves the rows within the block alone", async () => {
    const moved = await through(pivotAt("A12:B16"), shiftPivotLocation, INSERT);

    expect(moved).toContain('firstHeaderRow="1"');
    expect(moved).toContain('firstDataRow="1"');
  });
});

describe("withCacheRefreshedOnLoad", () => {
  it("asks for a rebuild on open", async () => {
    expect(await through(cacheOver("A1:C15"), (events) => withCacheRefreshedOnLoad(events), INSERT)).toContain(
      'refreshOnLoad="1"',
    );
  });

  it("keeps an existing setting from appearing twice", async () => {
    const already = '<pivotCacheDefinition xmlns="m" refreshOnLoad="0"></pivotCacheDefinition>';
    const refreshed = await through(already, (events) => withCacheRefreshedOnLoad(events), INSERT);

    expect(refreshed).toBe('<pivotCacheDefinition xmlns="m" refreshOnLoad="1"></pivotCacheDefinition>');
  });
});
