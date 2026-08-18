import type { CellInput } from "./cell-input";
import type { RowSource } from "./editor";
import type { NamedRegion } from "./named-region";

export interface RegionWrite {
  /**
   * Set when the region may run past its last row, to carry back how far it went.
   * Absent when the region is fixed and an extra row is an error.
   */
  readonly growth?: { lastRow: number };

  /** What to say when another row will not fit, where the plain message is not the useful one. */
  readonly whenFull?: string;
}

/**
 * The rows a region write puts on the sheet: the caller's values placed at the
 * region's first column, and every cell of the region the caller did not fill
 * blanked rather than left as it was.
 *
 * Throws if the caller offers more rows than the region holds, or a row wider
 * than it is, rather than writing past its edge into whatever the template keeps
 * there.
 */
export async function* regionRows(
  region: NamedRegion,
  rows: RowSource,
  write: RegionWrite = {},
): AsyncIterable<readonly (CellInput | undefined)[]> {
  const height = region.lastRow - region.firstRow + 1;
  const width = region.lastColumnIndex - region.firstColumnIndex + 1;
  const lead = Array.from<undefined>({ length: region.firstColumnIndex });
  let written = 0;

  for await (const values of rows) {
    if (written === height && write.growth === undefined) {
      throw new Error(write.whenFull ?? `The name "${region.name}" covers ${height} rows and was given more`);
    }
    if (values.length > width) {
      throw new Error(`The name "${region.name}" covers ${width} columns and was given a row of ${values.length}`);
    }

    yield [...lead, ...values, ...blanks(width - values.length)];
    written += 1;
  }

  // A row the caller did not reach is cleared, not left alone. What is sitting
  // there is the last run's data, formatted exactly like this run's, so leaving it
  // would read as a current number rather than as a leftover.
  for (; written < height; written += 1) {
    yield [...lead, ...blanks(width)];
  }

  // A region that grew only ever grew, so a run with fewer rows than last time
  // leaves the region the size it was and clears what it did not fill.
  if (write.growth !== undefined) {
    write.growth.lastRow = region.firstRow + written - 1;
  }
}

function blanks(count: number): null[] {
  return Array.from({ length: count }, () => null);
}
