import type { CellInput } from "./cell-input";
import type { RowSource } from "./editor";
import type { NamedRegion } from "./named-region";

/**
 * The rows to write into a region, and how many there are.
 *
 * The count has to be known before anything is written, because it decides how far
 * the sheet moves and the rows above the region go out first.
 */
export interface RegionRows {
  readonly count: number;

  /** Read once, as the region is written. */
  rows(): AsyncIterable<readonly (CellInput | undefined)[]>;
}

/**
 * `source` placed into `region`'s columns, without being counted, so nothing is
 * held. The caller learns how many rows there were by counting them as it writes.
 */

/**
 * Prepares `source` for writing into `region`, placing each row at the region's
 * first column and blanking every column of it the caller did not fill.
 *
 * No rows at all still produces one blank row. A region with nothing left in it
 * would take every reference to it down with it, and one surviving row keeps a
 * total written over the region alive.
 *
 * Throws on a row wider than the region, rather than writing into the column
 * beside it. Nothing moves sideways, so there is nowhere for it to go.
 */
export async function readRegionRows(region: NamedRegion, source: RowSource): Promise<RegionRows> {
  // An array already says how many rows it has, so nothing is held to find out and
  // the rows are read as the region is written, like every other write path.
  if (Array.isArray(source)) {
    return { count: Math.max(1, source.length), rows: () => placeEach(region, source) };
  }

  return held(region, source);
}

export async function* placeEach(
  region: NamedRegion,
  source: RowSource,
): AsyncIterable<readonly (CellInput | undefined)[]> {
  const width = widthOf(region);
  const row = scratchRow(region, width);
  let written = 0;

  for await (const values of source) {
    yield fill(row, region, width, values);
    written += 1;
  }

  if (written === 0) {
    yield fill(row, region, width, []);
  }
}

// Held in one flat run of values rather than as an array per row. A region holds
// every row it is given, and at a million rows the per-array overhead is most of
// what is held, where the values themselves are not.
async function held(region: NamedRegion, source: RowSource): Promise<RegionRows> {
  const width = widthOf(region);
  const values: (CellInput | undefined)[] = [];
  let count = 0;

  for await (const row of source) {
    checkWidth(region, width, row);

    for (let column = 0; column < width; column += 1) {
      values.push(column < row.length ? row[column] : null);
    }
    count += 1;
  }

  return {
    count: Math.max(1, count),
    rows: async function* (): AsyncIterable<readonly (CellInput | undefined)[]> {
      const row = scratchRow(region, width);

      for (let start = 0; start < values.length; start += width) {
        for (let column = 0; column < width; column += 1) {
          row[region.firstColumnIndex + column] = values[start + column];
        }

        yield row;
      }

      if (count === 0) {
        yield fill(row, region, width, []);
      }
    },
  };
}

function widthOf(region: NamedRegion): number {
  return region.lastColumnIndex - region.firstColumnIndex + 1;
}

function checkWidth(region: NamedRegion, width: number, values: readonly unknown[]): void {
  if (values.length > width) {
    throw new Error(`The name "${region.name}" covers ${width} columns and was given a row of ${values.length}`);
  }
}

// One row is filled over and over rather than allocated per row. A region can be
// given millions of rows, and at that size allocating one array each is most of
// what the writer spends. Whatever reads these turns each row into its own cells
// before asking for the next one, so no row outlives the call that yielded it.
//
// The caller's own array is never handed on either way, so a caller reusing one
// array for every row still works.
function scratchRow(region: NamedRegion, width: number): (CellInput | undefined)[] {
  return new Array<CellInput | undefined>(region.firstColumnIndex + width);
}

function fill(
  row: (CellInput | undefined)[],
  region: NamedRegion,
  width: number,
  values: readonly (CellInput | undefined)[],
): readonly (CellInput | undefined)[] {
  checkWidth(region, width, values);

  for (let column = 0; column < width; column += 1) {
    row[region.firstColumnIndex + column] = column < values.length ? values[column] : null;
  }

  return row;
}
