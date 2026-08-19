import { describe, expect, it } from "vitest";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { movedSheetEvent, movedSourceRow, pointsAtOrBelow } from "../../src/xlsx/shift-sheet";
import type { SourceRow } from "../../src/xlsx/write-sheet";
import type { XmlEvent } from "../../src/xml/xml-reader";

// Two rows appear at row 11, the last row of a region covering 9 to 11.
const INSERT: RowShift = { sheet: "Report", at: 11, by: 2 };

// Rows 10 and 11 go away.
const REMOVE: RowShift = { sheet: "Report", at: 10, by: -2 };

function rowAt(number: number, inner: readonly XmlEvent[] = []): SourceRow {
  return {
    number,
    attributes: { r: String(number), spans: "1:3" },
    cells: [{ columnIndex: 2, attributes: { r: `C${number}` }, inner }],
  };
}

const formula = (text: string): XmlEvent[] => [
  { type: "open", name: "f", attributes: {} },
  { type: "text", text },
  { type: "close", name: "f" },
];

describe("movedSourceRow", () => {
  it("renumbers the row and its cells together", () => {
    const moved = movedSourceRow(rowAt(12), INSERT);

    expect(moved?.number).toBe(14);
    expect(moved?.attributes["r"]).toBe("14");
    expect(moved?.cells[0]?.attributes["r"]).toBe("C14");
  });

  it("leaves a row above the move where it is", () => {
    expect(movedSourceRow(rowAt(9), INSERT)?.number).toBe(9);
  });

  it("has nowhere to put a row that was taken out", () => {
    expect(movedSourceRow(rowAt(10), REMOVE)).toBeUndefined();
  });

  it("moves the range a shared formula covers", () => {
    const shared: XmlEvent[] = [
      { type: "open", name: "f", attributes: { t: "shared", ref: "C12:C20", si: "0" } },
      { type: "text", text: "C11*2" },
      { type: "close", name: "f" },
    ];
    const moved = movedSourceRow(rowAt(12, shared), INSERT);

    expect(moved?.cells[0]?.inner[0]).toMatchObject({ attributes: { ref: "C14:C22" } });
    expect(moved?.cells[0]?.inner[1]).toEqual({ type: "text", text: "C13*2" });
  });

  it("rewrites a formula in a row that moved", () => {
    const moved = movedSourceRow(rowAt(12, formula("SUM(C9:C11)")), INSERT);

    expect(moved?.cells[0]?.inner[1]).toEqual({ type: "text", text: "SUM(C9:C13)" });
  });

  it("leaves the column spans alone, since no column moved", () => {
    expect(movedSourceRow(rowAt(12), INSERT)?.attributes["spans"]).toBe("1:3");
  });
});

describe("movedSheetEvent", () => {
  it("moves a merged range", () => {
    const event: XmlEvent = { type: "open", name: "mergeCell", attributes: { ref: "A12:C12" } };

    expect(movedSheetEvent(event, INSERT)).toEqual({
      type: "open",
      name: "mergeCell",
      attributes: { ref: "A14:C14" },
    });
  });

  it("leaves text and closing tags as they are", () => {
    const text: XmlEvent = { type: "text", text: "C12" };

    expect(movedSheetEvent(text, INSERT)).toBe(text);
  });

  it("moves every range of a conditional format that covers more than one", () => {
    const event: XmlEvent = { type: "open", name: "conditionalFormatting", attributes: { sqref: "B12 D12:D14" } };

    expect(movedSheetEvent(event, INSERT)).toMatchObject({ attributes: { sqref: "B14 D14:D16" } });
  });

  it("moves a data validation, a hyperlink and a filter", () => {
    const ranges: [string, string, string][] = [
      ["dataValidation", "sqref", "B12"],
      ["hyperlink", "ref", "A12"],
      ["autoFilter", "ref", "A8:C12"],
    ];

    for (const [name, attribute, value] of ranges) {
      const moved = movedSheetEvent({ type: "open", name, attributes: { [attribute]: value } }, INSERT);

      expect(moved).toMatchObject({ attributes: { [attribute]: value.replace("12", "14") } });
    }
  });

  it("moves the cell a frozen pane starts at", () => {
    const event: XmlEvent = { type: "open", name: "pane", attributes: { ySplit: "11", topLeftCell: "A12" } };

    expect(movedSheetEvent(event, INSERT)).toMatchObject({ attributes: { topLeftCell: "A14" } });
  });

  // A page break names the row it sits above by number rather than by reference.
  it("moves a page break", () => {
    const event: XmlEvent = { type: "open", name: "brk", attributes: { id: "12", max: "16383" } };

    expect(movedSheetEvent(event, INSERT)).toMatchObject({ attributes: { id: "14" } });
  });

  it("refuses an extension list", () => {
    const event: XmlEvent = { type: "open", name: "extLst", attributes: {} };

    expect(() => movedSheetEvent(event, INSERT)).toThrow("holds an extension list");
  });
});

describe("pointsAtOrBelow", () => {
  it("finds a formula reading a row at or below the one asked about", () => {
    expect(pointsAtOrBelow(formula("SUM(C9:C11)"), 9)).toBe(true);
  });

  it("is false when everything points above it", () => {
    expect(pointsAtOrBelow(formula("SUM(C1:C8)"), 9)).toBe(false);
  });

  it("finds a range on an element's attribute", () => {
    const event: XmlEvent = { type: "open", name: "mergeCell", attributes: { ref: "A1:C12" } };

    expect(pointsAtOrBelow([event], 9)).toBe(true);
  });

  it("ignores an attribute that counts columns rather than naming a row", () => {
    const event: XmlEvent = { type: "open", name: "row", attributes: { r: "2", spans: "1:30" } };

    expect(pointsAtOrBelow([event], 9)).toBe(false);
  });
});
