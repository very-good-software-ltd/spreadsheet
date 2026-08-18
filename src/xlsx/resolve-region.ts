import type { NamedRegion } from "../named-region";
import type { TableOnSheet } from "./read-tables";
import type { DefinedNameRef } from "./read-workbook";

/**
 * The region a name points at, for a worksheet asking about its own names or for
 * the workbook asking about the ones it holds itself.
 *
 * A worksheet sees its own names before the workbook-wide ones of the same
 * spelling, and never another worksheet's. The workbook sees only its own.
 *
 * Throws when there is no such name, when the name points at something that
 * cannot be written, saying which, or when it points at a different worksheet
 * than the one asking.
 */
export interface NamedThings {
  readonly definedNames: readonly DefinedNameRef[];
  readonly tables: readonly TableOnSheet[];
}

export interface ResolvedRegion {
  readonly region: NamedRegion;

  /** The table the name belonged to, absent when it was a defined name. */
  readonly table?: TableOnSheet;
}

export function resolveRegion(named: NamedThings, name: string, worksheet?: string): ResolvedRegion {
  const { definedNames, tables } = named;
  const found =
    (worksheet === undefined ? undefined : find(definedNames, name, worksheet)) ?? find(definedNames, name, undefined);

  if (found === undefined) {
    // Excel keeps tables and defined names in one namespace and will not let a
    // workbook hold both spellings, so which of the two is looked at first cannot
    // decide anything in a file it wrote.
    const table = tables.find((candidate) => sameName(candidate.name, name));
    if (table !== undefined) {
      return { region: dataRegionOf(table, worksheet), table };
    }

    throw new Error(
      worksheet === undefined
        ? `No name "${name}" in this workbook`
        : `No name "${name}" on worksheet "${worksheet}" or in this workbook`,
    );
  }

  if (found.target.kind === "unusable") {
    throw new Error(`The name "${found.name}" is ${found.target.reason}, so it cannot be written`);
  }

  if (worksheet !== undefined && found.target.sheet !== worksheet) {
    throw new Error(`The name "${found.name}" points at worksheet "${found.target.sheet}", not "${worksheet}"`);
  }

  const { sheet, firstRow, lastRow, firstColumnIndex, lastColumnIndex } = found.target;

  return { region: { name: found.name, sheet, firstRow, lastRow, firstColumnIndex, lastColumnIndex } };
}

// A table's extent covers its header row and its totals row, and neither is a
// place for a caller's data, so what is writable is the rows between them.
function dataRegionOf(table: TableOnSheet, worksheet: string | undefined): NamedRegion {
  if (worksheet !== undefined && table.sheet !== worksheet) {
    throw new Error(`The name "${table.name}" points at worksheet "${table.sheet}", not "${worksheet}"`);
  }

  const firstRow = table.firstRow + table.headerRowCount;
  const lastRow = table.lastRow - table.totalsRowCount;

  if (lastRow < firstRow) {
    throw new Error(`The table "${table.name}" has no rows between its header and its totals row`);
  }

  return {
    name: table.name,
    sheet: table.sheet,
    firstRow,
    lastRow,
    firstColumnIndex: table.firstColumnIndex,
    lastColumnIndex: table.lastColumnIndex,
  };
}

// Excel matches a name without regard to case, and its Name Manager will not
// define two that differ only in case, so a lookup does not either.
function find(
  definedNames: readonly DefinedNameRef[],
  name: string,
  scope: string | undefined,
): DefinedNameRef | undefined {
  return definedNames.find((defined) => defined.scope === scope && sameName(defined.name, name));
}

function sameName(one: string, other: string): boolean {
  return one.localeCompare(other, undefined, { sensitivity: "accent" }) === 0;
}
