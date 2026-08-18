import { describe, expect, it } from "vitest";
import type { DefinedNameRef } from "../../src/xlsx/read-workbook";
import { resolveRegion } from "../../src/xlsx/resolve-region";

function named(name: string, target: string, scope?: string): DefinedNameRef {
  const [sheet, area] = target.split("!");
  const corners = (area ?? "").split(":");
  const rows = corners.map((corner) => Number(corner.replaceAll(/[$A-Z]/g, "")));
  const columns = corners.map((corner) => corner.replaceAll(/[$0-9]/g, "").charCodeAt(0) - 65);

  return {
    name,
    scope,
    target: {
      kind: "region",
      sheet: sheet ?? "",
      firstRow: Math.min(...rows),
      lastRow: Math.max(...rows),
      firstColumnIndex: Math.min(...columns),
      lastColumnIndex: Math.max(...columns),
    },
  };
}

const DATA = named("Data", "Sheet1!$B$4:$D$20");

describe("resolveRegion", () => {
  it("finds a workbook-wide name and keeps where it points", () => {
    expect(resolveRegion([DATA], "Data")).toEqual({
      name: "Data",
      sheet: "Sheet1",
      firstRow: 4,
      lastRow: 20,
      firstColumnIndex: 1,
      lastColumnIndex: 3,
    });
  });

  it("finds a workbook-wide name from a worksheet that has none of its own", () => {
    expect(resolveRegion([DATA], "Data", "Sheet1")).toMatchObject({ firstRow: 4 });
  });

  it("prefers a worksheet's own name over the workbook-wide one of the same spelling", () => {
    const names = [DATA, named("Data", "Sheet1!$A$1:$A$2", "Sheet1")];

    expect(resolveRegion(names, "Data", "Sheet1")).toMatchObject({ firstRow: 1, lastRow: 2 });
  });

  it("does not let a worksheet's own name leak into the workbook-wide lookup", () => {
    const names = [named("Data", "Sheet1!$A$1:$A$2", "Sheet1")];

    expect(() => resolveRegion(names, "Data")).toThrow('No name "Data" in this workbook');
  });

  it("does not let one worksheet see another's name", () => {
    const names = [named("Data", "Sheet2!$A$1:$A$2", "Sheet2")];

    expect(() => resolveRegion(names, "Data", "Sheet1")).toThrow(
      'No name "Data" on worksheet "Sheet1" or in this workbook',
    );
  });

  // Excel's Name Manager will not define two names differing only in case.
  it("matches a name whatever case it is asked for in", () => {
    expect(resolveRegion([DATA], "DATA")).toMatchObject({ name: "Data" });
  });

  it("says what the name is when it points at nothing writable", () => {
    const names: DefinedNameRef[] = [
      { name: "Rate", scope: undefined, target: { kind: "unusable", reason: "not a range" } },
      { name: "Column", scope: undefined, target: { kind: "unusable", reason: "a whole column" } },
    ];

    expect(() => resolveRegion(names, "Rate")).toThrow('The name "Rate" is not a range, so it cannot be written');
    expect(() => resolveRegion(names, "Column")).toThrow(
      'The name "Column" is a whole column, so it cannot be written',
    );
  });

  it("refuses a name that points at a different worksheet than the one asking", () => {
    expect(() => resolveRegion([DATA], "Data", "Summary")).toThrow(
      'The name "Data" points at worksheet "Sheet1", not "Summary"',
    );
  });
});
