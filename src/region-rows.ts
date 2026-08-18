import type { CellInput } from "./cell-input";
import type { RowSource } from "./editor";
import type { NamedRegion } from "./named-region";

/**
 * The rows a region write puts on the sheet: the caller's values placed at the
 * region's first column, with every column of the region the caller did not fill
 * blanked rather than left as it was.
 *
 * There are exactly as many rows as the caller gave, except that no rows at all
 * still produces one blank row. A region with nothing left in it would take every
 * reference to it down with it, and one surviving row keeps a total written over
 * the region alive.
 *
 * Throws on a row wider than the region, rather than writing into the column
 * beside it. Nothing moves sideways, so there is nowhere for it to go.
 */
export async function* regionRows(
  region: NamedRegion,
  rows: RowSource,
): AsyncIterable<readonly (CellInput | undefined)[]> {
  const width = region.lastColumnIndex - region.firstColumnIndex + 1;
  const lead = Array.from<undefined>({ length: region.firstColumnIndex });
  let written = 0;

  for await (const values of rows) {
    if (values.length > width) {
      throw new Error(`The name "${region.name}" covers ${width} columns and was given a row of ${values.length}`);
    }

    yield [...lead, ...values, ...blanks(width - values.length)];
    written += 1;
  }

  if (written === 0) {
    yield [...lead, ...blanks(width)];
  }
}

function blanks(count: number): null[] {
  return Array.from({ length: count }, () => null);
}
