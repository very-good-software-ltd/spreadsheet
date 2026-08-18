import { columnIndexOf, rowNumberOf } from "../cell-reference";
import type { Region } from "../named-region";

export type DefinedNameTarget =
  | ({ readonly kind: "region" } & Region)
  | { readonly kind: "unusable"; readonly reason: string };

// Excel writes this in place of a reference whose sheet or cells were deleted. It
// can stand for the whole target or for either end of a range.
const BROKEN_REFERENCE = "#REF!";

const ABSOLUTE_CELL = /^\$[A-Z]+\$\d+$/;
const CELL = /^\$?[A-Z]+\$?\d+$/;
const COLUMN_ONLY = /^\$?[A-Z]+$/;
const ROW_ONLY = /^\$?\d+$/;

const unusable = (reason: string): DefinedNameTarget => ({ kind: "unusable", reason });

/**
 * The place a `definedName` points at, or the reason it points at nothing we can
 * write into.
 *
 * Only a single absolute cell range on one sheet of this workbook is usable. The
 * element also holds constants, formulas, multi-area ranges, whole columns and
 * rows, references into other workbooks, and references Excel broke on a delete,
 * each of which comes back as its own reason.
 */
export function parseDefinedName(target: string): DefinedNameTarget {
  const text = target.trim();

  if (text.includes(BROKEN_REFERENCE)) {
    return unusable("a broken reference");
  }
  if (text.startsWith("[")) {
    return unusable("a reference to another workbook");
  }

  const split = splitSheet(text);
  if (split === undefined) {
    return isMultiArea(text) ? unusable("a range covering more than one area") : unusable("not a range");
  }

  const { sheet, area } = split;
  if (area.includes(",")) {
    return unusable("a range covering more than one area");
  }
  if (area.includes("!")) {
    return unusable("a range spanning more than one sheet");
  }

  const corners = area.split(":");
  if (corners.length > 2) {
    return unusable("not a range");
  }

  if (corners.every((corner) => COLUMN_ONLY.test(corner))) {
    return unusable("a whole column");
  }
  if (corners.every((corner) => ROW_ONLY.test(corner))) {
    return unusable("a whole row");
  }

  if (!corners.every((corner) => ABSOLUTE_CELL.test(corner))) {
    // A reference without both dollars resolves against whichever cell is selected,
    // which is a notion of the Excel UI and has no fixed position in a file.
    return corners.every((corner) => CELL.test(corner))
      ? unusable("a relative reference, which has no fixed position")
      : unusable("not a range");
  }

  const cells = corners.map((corner) => corner.replaceAll("$", ""));
  const rows = cells.map(rowNumberOf);
  const columns = cells.map(columnIndexOf);

  return {
    kind: "region",
    sheet,
    firstRow: Math.min(...rows),
    lastRow: Math.max(...rows),
    firstColumnIndex: Math.min(...columns),
    lastColumnIndex: Math.max(...columns),
  };
}

// A sheet name holding a space, a comma or a bracket is quoted, and a quote inside
// it is written twice, the same escape a formula uses.
function splitSheet(text: string): { sheet: string; area: string } | undefined {
  if (!text.startsWith("'")) {
    const marker = text.indexOf("!");
    if (marker <= 0) {
      return undefined;
    }
    const sheet = text.slice(0, marker);
    return /[[\]:,]/.test(sheet) ? undefined : { sheet, area: text.slice(marker + 1) };
  }

  let position = 1;
  let sheet = "";
  while (position < text.length) {
    if (text[position] !== "'") {
      sheet += text[position];
      position += 1;
    } else if (text[position + 1] === "'") {
      sheet += "'";
      position += 2;
    } else {
      break;
    }
  }

  return text[position] === "'" && text[position + 1] === "!" ? { sheet, area: text.slice(position + 2) } : undefined;
}

function isMultiArea(text: string): boolean {
  const areas = text.split(",");
  return areas.length > 1 && areas.every((area) => splitSheet(area.trim()) !== undefined);
}
