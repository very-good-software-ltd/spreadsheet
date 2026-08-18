import type { NamedRegion } from "../named-region";
import type { RowShift } from "./shift-formula";

/**
 * How the sheet has to move for `region` to hold `rows` rows, or `undefined` when
 * it already holds exactly that many.
 *
 * Room is made at the region's last row rather than after it, because Excel
 * stretches a range only when rows appear strictly inside it. A total written
 * `=SUM(C9:C11)` over rows 9 to 11 then covers the new rows too, where inserting
 * below row 11 would leave it summing three of them.
 *
 * Rows are taken from the far end for the matching reason, so the rows kept are
 * the first ones and a range over the region closes up against what survives.
 */
export function shiftFor(region: NamedRegion, rows: number): RowShift | undefined {
  const height = region.lastRow - region.firstRow + 1;
  const by = rows - height;

  if (by === 0) {
    return undefined;
  }

  return {
    sheet: region.sheet,
    at: by > 0 ? region.lastRow : region.firstRow + rows,
    by,
  };
}
