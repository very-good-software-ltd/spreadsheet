import { describe, expect, it } from "vitest";
import { Formula, formula } from "../src/cell-input";

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
