import { describe, expect, it } from "vitest";
import { dateToSerial, serialToDate } from "../../src/xlsx/date";

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

describe("dateToSerial", () => {
  it("maps dates to their 1900-system serials", () => {
    expect(dateToSerial(new Date("1900-01-01T00:00:00.000Z"), false)).toBe(1);
    expect(dateToSerial(new Date("1900-03-01T00:00:00.000Z"), false)).toBe(61);
    expect(dateToSerial(new Date("1970-01-01T00:00:00.000Z"), false)).toBe(25569);
    expect(dateToSerial(new Date("2020-01-01T00:00:00.000Z"), false)).toBe(43831);
  });

  it("maps dates to their 1904-system serials", () => {
    expect(dateToSerial(new Date("1904-01-01T00:00:00.000Z"), true)).toBe(0);
    expect(dateToSerial(new Date("1904-01-02T00:00:00.000Z"), true)).toBe(1);
    // The two systems sit 1462 days apart, four years plus the 1904 leap day.
    expect(dateToSerial(new Date("1970-01-01T00:00:00.000Z"), true)).toBe(25569 - 1462);
  });

  it("picks the real serial where the phantom leap day makes two map to one date", () => {
    // Serial 60 is Excel's 1900-02-29, a day that never existed, so reading it
    // and reading 59 both give 1900-02-28. Writing that date has to pick, and
    // 59 is the one that means it.
    expect(serialToDate(60, false).toISOString()).toBe("1900-02-28T00:00:00.000Z");
    expect(serialToDate(59, false).toISOString()).toBe("1900-02-28T00:00:00.000Z");
    expect(dateToSerial(new Date("1900-02-28T00:00:00.000Z"), false)).toBe(59);
  });

  it("carries the time of day as the fraction of a day", () => {
    expect(dateToSerial(new Date("1970-01-01T12:00:00.000Z"), false)).toBe(25569.5);
    expect(dateToSerial(new Date("1970-01-01T06:00:00.000Z"), false)).toBe(25569.25);
  });

  it("round-trips every serial serialToDate can produce", () => {
    for (const serial of [1, 59, 61, 25569, 43831, 43831.5]) {
      expect(dateToSerial(serialToDate(serial, false), false)).toBe(serial);
      expect(dateToSerial(serialToDate(serial, true), true)).toBe(serial);
    }
  });

  it("reads the date in UTC, matching what serialToDate builds", () => {
    // A serial carries no timezone, so the conversion has to pick one, and
    // reading picked UTC. Local-midnight and UTC-midnight are different instants
    // and produce different serials, which is a real trap for callers.
    expect(dateToSerial(new Date(Date.UTC(2020, 0, 1)), false)).toBe(43831);
  });
});
