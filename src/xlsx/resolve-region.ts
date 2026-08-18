import type { NamedRegion } from "../named-region";
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
export function resolveRegion(definedNames: readonly DefinedNameRef[], name: string, worksheet?: string): NamedRegion {
  const found =
    (worksheet === undefined ? undefined : find(definedNames, name, worksheet)) ?? find(definedNames, name, undefined);

  if (found === undefined) {
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

  return { name: found.name, sheet, firstRow, lastRow, firstColumnIndex, lastColumnIndex };
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
