import { describe, expect, it } from "vitest";
import type { TableOnSheet } from "../../src/xlsx/read-tables";
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
    expect(resolveRegion({ definedNames: [DATA], tables: [] }, "Data")).toEqual({
      name: "Data",
      sheet: "Sheet1",
      firstRow: 4,
      lastRow: 20,
      firstColumnIndex: 1,
      lastColumnIndex: 3,
    });
  });

  it("finds a workbook-wide name from a worksheet that has none of its own", () => {
    expect(resolveRegion({ definedNames: [DATA], tables: [] }, "Data", "Sheet1")).toMatchObject({ firstRow: 4 });
  });

  it("prefers a worksheet's own name over the workbook-wide one of the same spelling", () => {
    const names = [DATA, named("Data", "Sheet1!$A$1:$A$2", "Sheet1")];

    expect(resolveRegion({ definedNames: names, tables: [] }, "Data", "Sheet1")).toMatchObject({
      firstRow: 1,
      lastRow: 2,
    });
  });

  it("does not let a worksheet's own name leak into the workbook-wide lookup", () => {
    const names = [named("Data", "Sheet1!$A$1:$A$2", "Sheet1")];

    expect(() => resolveRegion({ definedNames: names, tables: [] }, "Data")).toThrow('No name "Data" in this workbook');
  });

  it("does not let one worksheet see another's name", () => {
    const names = [named("Data", "Sheet2!$A$1:$A$2", "Sheet2")];

    expect(() => resolveRegion({ definedNames: names, tables: [] }, "Data", "Sheet1")).toThrow(
      'No name "Data" on worksheet "Sheet1" or in this workbook',
    );
  });

  // Excel's Name Manager will not define two names differing only in case.
  it("matches a name whatever case it is asked for in", () => {
    expect(resolveRegion({ definedNames: [DATA], tables: [] }, "DATA")).toMatchObject({ name: "Data" });
  });

  it("says what the name is when it points at nothing writable", () => {
    const names: DefinedNameRef[] = [
      { name: "Rate", scope: undefined, target: { kind: "unusable", reason: "not a range" } },
      { name: "Column", scope: undefined, target: { kind: "unusable", reason: "a whole column" } },
    ];

    expect(() => resolveRegion({ definedNames: names, tables: [] }, "Rate")).toThrow(
      'The name "Rate" is not a range, so it cannot be written',
    );
    expect(() => resolveRegion({ definedNames: names, tables: [] }, "Column")).toThrow(
      'The name "Column" is a whole column, so it cannot be written',
    );
  });

  it("refuses a name that points at a different worksheet than the one asking", () => {
    expect(() => resolveRegion({ definedNames: [DATA], tables: [] }, "Data", "Summary")).toThrow(
      'The name "Data" points at worksheet "Sheet1", not "Summary"',
    );
  });
});

function table(name: string, sheet: string, extent: string, rows: Partial<TableOnSheet> = {}): TableOnSheet {
  const [first, last] = extent.split(":");
  return {
    name,
    sheet,
    path: "xl/tables/table1.xml",
    firstRow: Number((first ?? "").replaceAll(/[A-Z]/g, "")),
    lastRow: Number((last ?? "").replaceAll(/[A-Z]/g, "")),
    firstColumnIndex: (first ?? "").charCodeAt(0) - 65,
    lastColumnIndex: (last ?? "").charCodeAt(0) - 65,
    headerRowCount: 1,
    totalsRowCount: 0,
    ...rows,
  };
}

describe("resolveRegion over tables", () => {
  // A table's extent covers its header and its totals row, and neither is a place
  // for a caller's data, so what comes back is the rows between them.
  it("gives a table's data rows, not its header or its totals row", () => {
    const tables = [table("Sales", "Sheet1", "B2:D9", { totalsRowCount: 1 })];

    expect(resolveRegion({ definedNames: [], tables }, "Sales")).toEqual({
      name: "Sales",
      sheet: "Sheet1",
      firstRow: 3,
      lastRow: 8,
      firstColumnIndex: 1,
      lastColumnIndex: 3,
    });
  });

  it("starts at the first row when the table declares no header", () => {
    const tables = [table("Sales", "Sheet1", "B2:D9", { headerRowCount: 0 })];

    expect(resolveRegion({ definedNames: [], tables }, "Sales")).toMatchObject({ firstRow: 2, lastRow: 9 });
  });

  it("finds a table from the worksheet it is on", () => {
    const tables = [table("Sales", "Sheet1", "B2:D9")];

    expect(resolveRegion({ definedNames: [], tables }, "Sales", "Sheet1")).toMatchObject({ firstRow: 3 });
  });

  it("refuses a table on a different worksheet than the one asking", () => {
    const tables = [table("Sales", "Sheet2", "B2:D9")];

    expect(() => resolveRegion({ definedNames: [], tables }, "Sales", "Sheet1")).toThrow(
      'The name "Sales" points at worksheet "Sheet2", not "Sheet1"',
    );
  });

  it("says a table has no room for data when it is only a header and a totals row", () => {
    const tables = [table("Sales", "Sheet1", "B2:D3", { totalsRowCount: 1 })];

    expect(() => resolveRegion({ definedNames: [], tables }, "Sales")).toThrow(
      'The table "Sales" has no rows between its header and its totals row',
    );
  });
});
