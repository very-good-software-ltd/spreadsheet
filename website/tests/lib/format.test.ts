import type { Cell } from "very-good-spreadsheet";
import { describe, expect, it } from "vitest";
import { columnLetter, formatBytes, formatCell, formatDate } from "~/lib/format";

function cell(value: Cell): Cell {
  return value;
}

describe("columnLetter", () => {
  it("maps zero-based indexes to spreadsheet letters", () => {
    expect(columnLetter(0)).toBe("A");
    expect(columnLetter(25)).toBe("Z");
    expect(columnLetter(26)).toBe("AA");
    expect(columnLetter(701)).toBe("ZZ");
  });
});

describe("formatBytes", () => {
  it("scales bytes to a readable unit", () => {
    expect(formatBytes(512)).toBe("512 B");
    expect(formatBytes(2048)).toBe("2 KB");
    expect(formatBytes(5 * 1024 * 1024)).toBe("5.0 MB");
  });
});

describe("formatDate", () => {
  it("shows a date on its own when there is no time", () => {
    expect(formatDate(new Date(Date.UTC(2020, 0, 15)))).toBe("2020-01-15");
  });

  it("appends the time when the date has one", () => {
    expect(formatDate(new Date(Date.UTC(2020, 0, 15, 13, 45)))).toBe("2020-01-15 13:45");
  });
});

describe("formatCell", () => {
  it("right-aligns numbers", () => {
    expect(formatCell(cell({ ref: "A1", columnIndex: 0, type: "number", value: 42 }))).toEqual({
      text: "42",
      align: "text-right",
      tone: "",
    });
  });

  it("right-aligns and formats dates", () => {
    const result = formatCell(
      cell({ ref: "A1", columnIndex: 0, type: "date", value: new Date(Date.UTC(2020, 0, 15)) }),
    );
    expect(result.text).toBe("2020-01-15");
    expect(result.align).toBe("text-right");
  });

  it("shows booleans as TRUE and FALSE", () => {
    expect(formatCell(cell({ ref: "A1", columnIndex: 0, type: "boolean", value: true })).text).toBe("TRUE");
    expect(formatCell(cell({ ref: "A1", columnIndex: 0, type: "boolean", value: false })).text).toBe("FALSE");
  });

  it("tints errors red", () => {
    const result = formatCell(cell({ ref: "A1", columnIndex: 0, type: "error", value: "#DIV/0!" }));
    expect(result.text).toBe("#DIV/0!");
    expect(result.tone).toContain("red");
  });

  it("leaves strings left-aligned and untinted", () => {
    expect(formatCell(cell({ ref: "A1", columnIndex: 0, type: "string", value: "hello" }))).toEqual({
      text: "hello",
      align: "text-left",
      tone: "",
    });
  });
});
