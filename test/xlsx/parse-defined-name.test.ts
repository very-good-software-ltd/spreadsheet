import { describe, expect, it } from "vitest";
import { parseDefinedName } from "../../src/xlsx/parse-defined-name";

describe("parseDefinedName", () => {
  it("reads an absolute range on one sheet", () => {
    expect(parseDefinedName("Sheet1!$B$4:$D$20")).toEqual({
      kind: "region",
      sheet: "Sheet1",
      firstRow: 4,
      lastRow: 20,
      firstColumnIndex: 1,
      lastColumnIndex: 3,
    });
  });

  it("reads a single cell as a region one cell wide and tall", () => {
    expect(parseDefinedName("Sheet1!$B$4")).toEqual({
      kind: "region",
      sheet: "Sheet1",
      firstRow: 4,
      lastRow: 4,
      firstColumnIndex: 1,
      lastColumnIndex: 1,
    });
  });

  it("unquotes a sheet name that needed quoting", () => {
    expect(parseDefinedName("'My Sheet'!$A$1:$A$2")).toMatchObject({ sheet: "My Sheet" });
  });

  // A quote inside a sheet name is written twice, the same escape Excel uses in a formula.
  it("unescapes a doubled quote in a sheet name", () => {
    expect(parseDefinedName("'Bob''s Sheet'!$A$1")).toMatchObject({ sheet: "Bob's Sheet" });
  });

  it("orders the corners, so a range written bottom-right first still reads low to high", () => {
    expect(parseDefinedName("Sheet1!$D$20:$B$4")).toMatchObject({
      firstRow: 4,
      lastRow: 20,
      firstColumnIndex: 1,
      lastColumnIndex: 3,
    });
  });

  it("refuses a name that is a constant or a formula rather than a place", () => {
    expect(parseDefinedName("42")).toEqual({ kind: "unusable", reason: "not a range" });
    expect(parseDefinedName("SUM(Sheet1!$A$1:$A$9)")).toEqual({
      kind: "unusable",
      reason: "not a range",
    });
  });

  it("refuses a range covering more than one area", () => {
    expect(parseDefinedName("Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2")).toEqual({
      kind: "unusable",
      reason: "a range covering more than one area",
    });
  });

  // Without the dollars a reference resolves against whichever cell is selected,
  // which is a UI notion with no meaning outside Excel.
  it("refuses a reference that is not fully absolute", () => {
    const reason = "a relative reference, which has no fixed position";
    expect(parseDefinedName("Sheet1!B4:D20")).toEqual({ kind: "unusable", reason });
    expect(parseDefinedName("Sheet1!$B4:$D20")).toEqual({ kind: "unusable", reason });
    expect(parseDefinedName("Sheet1!B$4:D$20")).toEqual({ kind: "unusable", reason });
  });

  it("refuses a whole column or a whole row", () => {
    expect(parseDefinedName("Sheet1!$B:$B")).toEqual({
      kind: "unusable",
      reason: "a whole column",
    });
    expect(parseDefinedName("Sheet1!$4:$9")).toEqual({ kind: "unusable", reason: "a whole row" });
  });

  it("refuses a reference into another workbook", () => {
    expect(parseDefinedName("[1]Sheet1!$A$1")).toEqual({
      kind: "unusable",
      reason: "a reference to another workbook",
    });
  });

  // Excel leaves this behind when the sheet or the cells a name pointed at are deleted.
  it("refuses a reference broken by a deletion", () => {
    const reason = "a broken reference";
    expect(parseDefinedName("#REF!")).toEqual({ kind: "unusable", reason });
    expect(parseDefinedName("Sheet1!#REF!")).toEqual({ kind: "unusable", reason });
    expect(parseDefinedName("#REF!$A$1")).toEqual({ kind: "unusable", reason });
  });

  it("refuses a range spanning two sheets", () => {
    expect(parseDefinedName("Sheet1!$A$1:Sheet2!$A$1")).toEqual({
      kind: "unusable",
      reason: "a range spanning more than one sheet",
    });
  });
});
