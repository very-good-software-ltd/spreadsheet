import { describe, expect, it } from "vitest";
import type { NamedRegion } from "../../src/named-region";
import { shiftFor } from "../../src/xlsx/region-shift";

// Rows 9 to 11 of Report, three rows tall.
const REGION: NamedRegion = {
  name: "Data",
  sheet: "Report",
  firstRow: 9,
  lastRow: 11,
  firstColumnIndex: 1,
  lastColumnIndex: 3,
};

describe("shiftFor", () => {
  it("does not move the sheet when the rows already fit", () => {
    expect(shiftFor(REGION, 3)).toBeUndefined();
  });

  // Excel stretches a range only when rows appear strictly inside it, so making
  // room at row 11 leaves =SUM(C9:C11) covering all five rows. Making it at row 12
  // would leave that total summing three of them.
  it("makes room at the last row so a total over the region stretches with it", () => {
    expect(shiftFor(REGION, 5)).toEqual({ sheet: "Report", at: 11, by: 2 });
  });

  it("takes rows from the far end, so the first row and its formatting stay", () => {
    expect(shiftFor(REGION, 1)).toEqual({ sheet: "Report", at: 10, by: -2 });
  });

  it("takes only the rows it has to", () => {
    expect(shiftFor(REGION, 2)).toEqual({ sheet: "Report", at: 11, by: -1 });
  });
});
