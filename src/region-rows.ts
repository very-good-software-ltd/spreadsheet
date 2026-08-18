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
  let written = 0;

  for await (const values of rows) {
    if (values.length > width) {
      throw new Error(`The name "${region.name}" covers ${width} columns and was given a row of ${values.length}`);
    }

    yield placed(region, width, values);
    written += 1;
  }

  if (written === 0) {
    yield placed(region, width, []);
  }
}

// Built in one allocation rather than by spreading three arrays together. Every
// row a region write is given is held until the file is written, so what this
// costs per row is what it costs a million times over.
//
// The caller's own array is never handed on, so a caller reusing one array for
// every row still works.
function placed(
  region: NamedRegion,
  width: number,
  values: readonly (CellInput | undefined)[],
): readonly (CellInput | undefined)[] {
  const row = new Array<CellInput | undefined>(region.firstColumnIndex + width);

  for (let column = 0; column < width; column += 1) {
    row[region.firstColumnIndex + column] = column < values.length ? values[column] : null;
  }

  return row;
}
