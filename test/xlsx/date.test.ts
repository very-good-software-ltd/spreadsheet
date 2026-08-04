import { describe, expect, it } from "vitest";
import { serialToDate } from "../../src/xlsx/date";

describe("serialToDate", () => {
  it("maps 1900-system serials, including dates before the phantom leap day", () => {
    expect(serialToDate(1, false).toISOString()).toBe("1900-01-01T00:00:00.000Z");
    expect(serialToDate(59, false).toISOString()).toBe("1900-02-28T00:00:00.000Z");
    expect(serialToDate(61, false).toISOString()).toBe("1900-03-01T00:00:00.000Z");
    expect(serialToDate(43831, false).toISOString()).toBe("2020-01-01T00:00:00.000Z");
  });

  it("maps 1904-system serials", () => {
    expect(serialToDate(0, true).toISOString()).toBe("1904-01-01T00:00:00.000Z");
    expect(serialToDate(1, true).toISOString()).toBe("1904-01-02T00:00:00.000Z");
  });
});
