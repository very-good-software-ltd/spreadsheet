import { describe, expect, it } from "vitest";
import { shiftDrawingAnchors } from "../../src/xlsx/shift-drawing";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { writeXmlEvent } from "../../src/xml/write-xml";

function streamOf(text: string): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(new TextEncoder().encode(text));
      controller.close();
    },
  });
}

async function shifted(drawing: string, shift: RowShift): Promise<string> {
  const events = shiftDrawingAnchors(new SaxesXmlReader().read(streamOf(drawing)), shift);
  let out = "";
  for await (const batch of events) {
    for (const event of batch) {
      out += writeXmlEvent(event);
    }
  }
  return out;
}

// A drawing counts rows from zero, so these are sheet rows 10 and 15.
function twoCell(from: number, to: number): string {
  return `<xdr:wsDr xmlns:xdr="x"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to></xdr:twoCellAnchor></xdr:wsDr>`;
}

function oneCell(from: number): string {
  return `<xdr:wsDr xmlns:xdr="x"><xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="476250" cy="476250"/></xdr:oneCellAnchor></xdr:wsDr>`;
}

const rowsOf = (drawing: string): number[] =>
  [...drawing.matchAll(/<xdr:row>(\d+)<\/xdr:row>/g)].map((match) => Number(match[1]));

// Two rows appear at sheet row 11.
const INSERT: RowShift = { sheet: "Report", at: 11, by: 2 };

// Sheet rows 10 and 11 go away.
const REMOVE: RowShift = { sheet: "Report", at: 10, by: -2 };

describe("shiftDrawingAnchors", () => {
  it("moves a shape anchored below the rows down by as many", async () => {
    expect(rowsOf(await shifted(twoCell(11, 15), INSERT))).toEqual([13, 17]);
  });

  it("leaves a shape anchored above the rows alone", async () => {
    expect(rowsOf(await shifted(twoCell(1, 5), INSERT))).toEqual([1, 5]);
  });

  // Sheet rows 10 to 16 with two rows appearing at 11: the top stays and the
  // bottom moves, which is what stretching means for a shape spanning the change.
  it("stretches a shape the rows appeared inside", async () => {
    expect(rowsOf(await shifted(twoCell(9, 15), INSERT))).toEqual([9, 17]);
  });

  it("pulls a shape below the removed rows up", async () => {
    expect(rowsOf(await shifted(twoCell(11, 15), REMOVE))).toEqual([9, 13]);
  });

  it("closes a shape up against rows that went away underneath its top", async () => {
    expect(rowsOf(await shifted(twoCell(8, 10), REMOVE))).toEqual([8, 8]);
  });

  it("moves a shape anchored to one cell", async () => {
    expect(rowsOf(await shifted(oneCell(11), INSERT))).toEqual([13]);
  });

  // Nothing is left for it to hang from, and quietly dropping a chart is worse
  // than saying so.
  it("refuses a shape standing entirely on rows that went away", async () => {
    await expect(shifted(twoCell(9, 10), REMOVE)).rejects.toThrow(
      "a drawing anchored only to rows that are being taken out",
    );
  });

  it("refuses a one cell shape whose only anchor went away", async () => {
    await expect(shifted(oneCell(9), REMOVE)).rejects.toThrow(
      "a drawing anchored only to rows that are being taken out",
    );
  });

  // An absolute anchor is placed in absolute units and names no row at all.
  it("leaves an absolutely positioned shape alone", async () => {
    const drawing = `<xdr:wsDr xmlns:xdr="x"><xdr:absoluteAnchor><xdr:pos x="0" y="100"/><xdr:ext cx="1" cy="1"/></xdr:absoluteAnchor></xdr:wsDr>`;

    expect(await shifted(drawing, INSERT)).toContain(`<xdr:pos x="0" y="100">`);
  });
});
