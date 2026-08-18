import type { XmlEvent } from "../xml/xml-reader";
import { movedRow, type RowShift } from "./shift-formula";

const Element = {
  OneCellAnchor: "xdr:oneCellAnchor",
  TwoCellAnchor: "xdr:twoCellAnchor",
  From: "xdr:from",
  To: "xdr:to",
  Row: "xdr:row",
} as const;

const NOTHING_TO_HANG_FROM = "This worksheet has a drawing anchored only to rows that are being taken out";

/**
 * A drawing's events with every shape anchored at or below the moved rows moved
 * with them.
 *
 * A drawing counts rows from zero where a sheet counts from one, so the numbers
 * here are one less than the ones everywhere else.
 *
 * An absolute anchor names no row and is left alone, which matches Excel: a shape
 * placed in absolute units does not move when rows do.
 *
 * Throws when a shape stands only on rows that are going away, since there is
 * nothing left to hang it from and dropping a chart in silence is worse than
 * refusing the file.
 */
export async function* shiftDrawingAnchors(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  let corner: "from" | "to" | undefined;
  let inRow = false;
  let topWentAway = false;

  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      if (event.type === "open") {
        if (event.name === Element.OneCellAnchor || event.name === Element.TwoCellAnchor) {
          topWentAway = false;
        }
        corner = event.name === Element.From ? "from" : event.name === Element.To ? "to" : corner;
        inRow = event.name === Element.Row;
      } else if (event.type === "close") {
        // A shape anchored to one cell hangs from its top alone, so losing that is
        // enough on its own. One anchored to two survives while either corner does.
        if (event.name === Element.OneCellAnchor && topWentAway) {
          throw new Error(NOTHING_TO_HANG_FROM);
        }
        if (event.name === Element.From || event.name === Element.To) {
          corner = undefined;
        }
        inRow = false;
      } else if (inRow && corner !== undefined) {
        const row = movedRow(Number(event.text) + 1, shift);

        if (row === undefined && corner === "from") {
          topWentAway = true;
        } else if (row === undefined && topWentAway) {
          throw new Error(NOTHING_TO_HANG_FROM);
        }

        moved.push({ ...event, text: String(anchorRow(row, corner, shift)) });
        continue;
      }

      moved.push(event);
    }

    yield moved;
  }
}

// A corner left standing on nothing closes up against the rows that went. The top
// lands on the first row still there, and the bottom on the last one before them.
function anchorRow(row: number | undefined, corner: "from" | "to", shift: RowShift): number {
  if (row !== undefined) {
    return row - 1;
  }

  return Math.max(0, (corner === "from" ? shift.at : shift.at - 1) - 1);
}
