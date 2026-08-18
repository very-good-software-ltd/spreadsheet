import { describe, expect, it } from "vitest";
import { shiftFormula } from "../../src/xlsx/shift-formula";

// Two rows appear at row 11, which is inside a region of rows 9 to 11.
const INSERT = { sheet: "Sheet1", at: 11, by: 2 };

// Rows 10 and 11 go away, leaving row 9 as the whole of the region.
const REMOVE = { sheet: "Sheet1", at: 10, by: -2 };

const shifted = (text: string, shift = INSERT, onSheet = "Sheet1") => shiftFormula(text, shift, onSheet);

describe("shiftFormula when rows are inserted", () => {
  // Excel stretches a range only when rows appear strictly inside it, which is why
  // room is made at the region's last row rather than after it.
  it("stretches a range the rows appeared inside", () => {
    expect(shifted("SUM(C9:C11)")).toBe("SUM(C9:C13)");
  });

  it("moves a reference below the rows down", () => {
    expect(shifted("C12+D12")).toBe("C14+D14");
  });

  it("leaves a reference above the rows alone", () => {
    expect(shifted("SUM(C1:C9)")).toBe("SUM(C1:C9)");
  });

  it("moves an absolute reference, since a dollar fixes the row against copying and not against this", () => {
    expect(shifted("$C$12")).toBe("$C$14");
  });

  it("moves a reference that names the sheet the rows moved on", () => {
    expect(shifted("Sheet1!C12")).toBe("Sheet1!C14");
  });

  it("leaves a reference to another sheet alone", () => {
    expect(shifted("Sheet2!C12")).toBe("Sheet2!C12");
  });

  it("leaves an unqualified reference alone when the formula is on another sheet", () => {
    expect(shifted("C12", INSERT, "Summary")).toBe("C12");
  });

  it("moves a reference whose sheet name had to be quoted", () => {
    expect(shifted("'My Sheet'!C12", { sheet: "My Sheet", at: 11, by: 2 }, "Other")).toBe("'My Sheet'!C14");
  });
});

describe("shiftFormula when rows are removed", () => {
  it("pulls a reference below the removed rows up", () => {
    expect(shifted("C12", REMOVE)).toBe("C10");
  });

  it("shrinks a range whose end was removed, rather than breaking it", () => {
    expect(shifted("SUM(C9:C11)", REMOVE)).toBe("SUM(C9:C9)");
  });

  it("shrinks a range whose start was removed", () => {
    expect(shifted("SUM(C11:C14)", REMOVE)).toBe("SUM(C10:C12)");
  });

  it("shrinks a range the removed rows sat inside", () => {
    expect(shifted("SUM(C9:C14)", REMOVE)).toBe("SUM(C9:C12)");
  });

  // Excel has nowhere left to point, and says so in the formula itself.
  it("breaks a single reference to a row that went away", () => {
    expect(shifted("C10", REMOVE)).toBe("#REF!");
  });

  it("breaks a range that was entirely inside the removed rows", () => {
    expect(shifted("SUM(C10:C11)", REMOVE)).toBe("SUM(#REF!)");
  });
});

describe("shiftFormula on things that only look like references", () => {
  const AT_ROW_ONE = { sheet: "Sheet1", at: 1, by: 2 };

  // A function name can end in digits, and LOG10 reads as column LOG row 10 to
  // anything matching on shape alone. A reference is never followed by a bracket.
  it("leaves a function name that reads like a reference alone", () => {
    expect(shifted("LOG10(A1)", AT_ROW_ONE)).toBe("LOG10(A3)");
  });

  it("leaves text inside quotes alone", () => {
    expect(shifted('IF(A1>0,"C12","")', AT_ROW_ONE)).toBe('IF(A3>0,"C12","")');
  });

  it("leaves a table's structured reference alone", () => {
    expect(shifted("SUM(Sales[Amount])", AT_ROW_ONE)).toBe("SUM(Sales[Amount])");
  });

  it("leaves a defined name alone", () => {
    expect(shifted("TaxRate*A1", AT_ROW_ONE)).toBe("TaxRate*A3");
  });

  it("leaves a name that merely starts like a reference alone", () => {
    expect(shifted("A1_TOTAL", AT_ROW_ONE)).toBe("A1_TOTAL");
  });

  it("leaves a whole column alone, since no row of it moved", () => {
    expect(shifted("SUM(C:C)", AT_ROW_ONE)).toBe("SUM(C:C)");
  });

  it("leaves a reference into another workbook alone", () => {
    expect(shifted("[1]Sheet1!C12")).toBe("[1]Sheet1!C12");
  });
});

describe("shiftFormula on what it will not touch", () => {
  it("refuses a reference across a range of sheets", () => {
    expect(() => shifted("SUM(Sheet1:Sheet3!C12)")).toThrow(
      "refers to a range of sheets, which cannot be moved with confidence",
    );
  });

  it("refuses a whole row reference", () => {
    expect(() => shifted("SUM($12:$12)")).toThrow("refers to whole rows, which cannot be moved with confidence");
  });
});
