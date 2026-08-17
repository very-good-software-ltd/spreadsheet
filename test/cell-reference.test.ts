import { describe, expect, it } from "vitest";
import { cellReference, columnIndexOf, columnLetters, rowNumberOf } from "../src/cell-reference";

describe("columnIndexOf", () => {
  it("converts the letters to a zero-based index", () => {
    expect(columnIndexOf("A1")).toBe(0);
    expect(columnIndexOf("B1")).toBe(1);
    expect(columnIndexOf("Z1")).toBe(25);
    expect(columnIndexOf("AA1")).toBe(26);
    expect(columnIndexOf("AB1")).toBe(27);
    expect(columnIndexOf("XFD1")).toBe(16383);
  });

  it("ignores the row part, however many digits it has", () => {
    expect(columnIndexOf("C3")).toBe(2);
    expect(columnIndexOf("C1048576")).toBe(2);
  });
});

describe("columnLetters", () => {
  it("converts a zero-based index to the letters", () => {
    expect(columnLetters(0)).toBe("A");
    expect(columnLetters(25)).toBe("Z");
    expect(columnLetters(26)).toBe("AA");
    expect(columnLetters(27)).toBe("AB");
    expect(columnLetters(16383)).toBe("XFD");
  });

  it("is the inverse of columnIndexOf", () => {
    for (const index of [0, 1, 25, 26, 51, 52, 701, 702, 16383]) {
      expect(columnIndexOf(`${columnLetters(index)}1`)).toBe(index);
    }
  });
});

describe("rowNumberOf", () => {
  it("reads the one-based row number", () => {
    expect(rowNumberOf("A1")).toBe(1);
    expect(rowNumberOf("XFD1048576")).toBe(1048576);
  });
});

describe("cellReference", () => {
  it("builds a reference from a one-based row and a zero-based column", () => {
    expect(cellReference(1, 0)).toBe("A1");
    expect(cellReference(3, 2)).toBe("C3");
    expect(cellReference(1048576, 16383)).toBe("XFD1048576");
  });

  it("round-trips", () => {
    expect(columnIndexOf(cellReference(7, 28))).toBe(28);
    expect(rowNumberOf(cellReference(7, 28))).toBe(7);
  });
});
