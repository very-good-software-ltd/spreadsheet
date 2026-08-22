import { describe, expect, it } from "vitest";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { shiftVmlAnchors } from "../../src/xlsx/shift-vml";
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

async function shifted(vml: string, shift: RowShift): Promise<string> {
  const events = shiftVmlAnchors(new SaxesXmlReader().read(streamOf(vml)), shift);
  let out = "";
  for await (const batch of events) {
    for (const event of batch) {
      out += writeXmlEvent(event);
    }
  }
  return out;
}

// VML counts rows from zero, so a note on sheet row 12 carries `<x:Row>11</x:Row>`
// and a box drawn over sheet rows 11 to 15 anchors at 10 and 14.
function note(cellRow: number, boxTop: number, boxBottom: number): string {
  return `<v:shape id="_x0000_s1025" type="#_x0000_t202"><v:textbox><div/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:Anchor>2, 6, ${boxTop - 1}, 14, 4, 2, ${boxBottom - 1}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${cellRow - 1}</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>`;
}

function drawing(...shapes: readonly string[]): string {
  return `<xml xmlns:v="v" xmlns:o="o" xmlns:x="x"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202"><v:stroke joinstyle="miter"/></v:shapetype>${shapes.join("")}</xml>`;
}

// Read back as sheet rows, so the numbers here match every other row number.
const cellRowsOf = (vml: string): number[] =>
  [...vml.matchAll(/<x:Row>(\d+)<\/x:Row>/g)].map((match) => Number(match[1]) + 1);

const boxRowsOf = (vml: string): number[] =>
  [...vml.matchAll(/<x:Anchor>([^<]+)<\/x:Anchor>/g)].flatMap((match) => {
    const parts = (match[1] ?? "").split(",").map((part) => Number(part) + 1);
    return [parts[2] ?? 0, parts[6] ?? 0];
  });

// Two rows appear at row 11.
const INSERT: RowShift = { sheet: "Report", at: 11, by: 2 };

// Rows 10 and 11 go away.
const REMOVE: RowShift = { sheet: "Report", at: 10, by: -2 };

describe("shiftVmlAnchors", () => {
  it("moves the cell a note below the rows is attached to", async () => {
    expect(cellRowsOf(await shifted(drawing(note(12, 11, 15)), INSERT))).toEqual([14]);
  });

  it("moves the box a note below the rows is drawn in", async () => {
    expect(boxRowsOf(await shifted(drawing(note(12, 11, 15)), INSERT))).toEqual([13, 17]);
  });

  it("leaves a note above the rows alone", async () => {
    const moved = await shifted(drawing(note(3, 2, 6)), INSERT);

    expect(cellRowsOf(moved)).toEqual([3]);
    expect(boxRowsOf(moved)).toEqual([2, 6]);
  });

  it("pulls a note below the removed rows up", async () => {
    const moved = await shifted(drawing(note(12, 12, 15)), REMOVE);

    expect(cellRowsOf(moved)).toEqual([10]);
    expect(boxRowsOf(moved)).toEqual([10, 13]);
  });

  it("drops a note whose cell went away", async () => {
    expect(await shifted(drawing(note(10, 9, 13)), REMOVE)).not.toContain("x:ClientData");
  });

  it("keeps the notes either side of one that went away", async () => {
    const moved = await shifted(drawing(note(3, 2, 6), note(10, 9, 13), note(12, 12, 15)), REMOVE);

    expect(cellRowsOf(moved)).toEqual([3, 10]);
  });

  it("closes the box up against the rows its top stood on", async () => {
    expect(boxRowsOf(await shifted(drawing(note(12, 10, 15)), REMOVE))).toEqual([10, 13]);
  });

  it("leaves the shape layout and shape type alone", async () => {
    const moved = await shifted(drawing(note(12, 11, 15)), INSERT);

    expect(moved).toContain('<o:idmap v:ext="edit" data="1">');
    expect(moved).toContain('<v:shapetype id="_x0000_t202">');
  });

  it("keeps a shape that has no cell of its own", async () => {
    const shape = '<v:shape id="_x0000_s2000"><x:ClientData ObjectType="Button"/></v:shape>';

    expect(await shifted(drawing(shape), INSERT)).toContain('ObjectType="Button"');
  });
});
