import { describe, expect, it } from "vitest";
import { date, Formula, formula, SpreadsheetDate } from "../src/cell-input";

describe("formula", () => {
  it("keeps the text as stored, without a leading equals", () => {
    expect(formula("SUM(A1:B1)").text).toBe("SUM(A1:B1)");
  });

  it("strips a leading equals, which is how a spreadsheet shows a formula but not how a file stores one", () => {
    expect(formula("=SUM(A1:B1)").text).toBe("SUM(A1:B1)");
  });

  it("trims surrounding whitespace before looking for the equals", () => {
    expect(formula("  =SUM(A1:B1)  ").text).toBe("SUM(A1:B1)");
  });

  it("strips only the first equals, so a comparison survives", () => {
    expect(formula("=A1=B1").text).toBe("A1=B1");
  });

  it("rejects an empty formula at the call rather than writing a meaningless cell", () => {
    expect(() => formula("")).toThrow(/empty/i);
    expect(() => formula("   ")).toThrow(/empty/i);
    expect(() => formula("=")).toThrow(/empty/i);
  });

  it("produces a value a caller cannot mistake for a string", () => {
    expect(formula("SUM(A1:B1)")).toBeInstanceOf(Formula);
  });
});

describe("date", () => {
  function iso(built: SpreadsheetDate): string {
    return built.value.toISOString();
  }

  it("reads a plain ISO date", () => {
    expect(iso(date("2026-03-01"))).toBe("2026-03-01T00:00:00.000Z");
  });

  it("reads a date and time with no zone as the values it names", () => {
    expect(iso(date("2026-03-01T14:30"))).toBe("2026-03-01T14:30:00.000Z");
  });

  it("reads a date and time with a zone as the instant it names", () => {
    expect(iso(date("2026-03-01T14:30:00Z"))).toBe("2026-03-01T14:30:00.000Z");
    expect(iso(date("2026-03-01T14:30:00+02:00"))).toBe("2026-03-01T12:30:00.000Z");
  });

  it("builds from parts, with a month from 1 to 12 rather than JavaScript's 0 to 11", () => {
    expect(iso(date(2026, 3, 1))).toBe("2026-03-01T00:00:00.000Z");
    expect(iso(date(2026, 12, 31))).toBe("2026-12-31T00:00:00.000Z");
  });

  it("takes a time with the parts", () => {
    expect(iso(date(2026, 3, 1, 14, 30, 15))).toBe("2026-03-01T14:30:15.000Z");
  });

  it("agrees with itself either way round", () => {
    expect(iso(date(2026, 3, 1, 14, 30))).toBe(iso(date("2026-03-01T14:30")));
  });

  it("produces a value the writer can turn into a serial", () => {
    expect(date("2026-03-01")).toBeInstanceOf(SpreadsheetDate);
  });

  it("rejects text that is not a date", () => {
    expect(() => date("not a date")).toThrow(/not a date/i);
    expect(() => date("")).toThrow(/not a date/i);
  });

  it("rejects parts that do not name a real day, rather than rolling them over", () => {
    expect(() => date(2026, 13, 1)).toThrow(/not a date/i);
    expect(() => date(2026, 0, 1)).toThrow(/not a date/i);
    expect(() => date(2026, 2, 30)).toThrow(/not a date/i);
    expect(() => date(2026, 4, 31)).toThrow(/not a date/i);
    expect(() => date(2026, 3, 1, 25)).toThrow(/not a date/i);
  });

  it("accepts a leap day in a leap year and rejects it otherwise", () => {
    expect(iso(date(2024, 2, 29))).toBe("2024-02-29T00:00:00.000Z");
    expect(() => date(2026, 2, 29)).toThrow(/not a date/i);
  });

  it("rejects parts that are not whole numbers", () => {
    expect(() => date(2026, 3, 1.5)).toThrow(/whole number/i);
  });

  // The reason this constructor exists rather than accepting a Date: the two ways
  // of building the same calendar day are different instants, and only one of them
  // is the day the caller meant.
  it("gives the caller a way to say which calendar day they mean", () => {
    const utcMidnight = new Date("2026-03-01T00:00:00.000Z");

    expect(iso(date(utcMidnight.toISOString()))).toBe("2026-03-01T00:00:00.000Z");
    expect(iso(date(2026, 3, 1))).toBe("2026-03-01T00:00:00.000Z");
  });
});
