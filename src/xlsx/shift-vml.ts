import type { XmlEvent } from "../xml/xml-reader";
import { movedRow, type RowShift } from "./shift-formula";

const Element = {
  Shape: "v:shape",
  Anchor: "x:Anchor",
  Row: "x:Row",
} as const;

// An anchor is eight numbers, a column, a column offset, a row and a row offset for
// each corner of the box, so the rows are the third and the seventh.
const ANCHOR_CORNERS: ReadonlyMap<number, "from" | "to"> = new Map([
  [2, "from"],
  [6, "to"],
]);

const NUMBER = /-?\d+/;

/**
 * A VML part's events with every shape at or below the moved rows moved with them.
 *
 * VML counts rows from zero where a sheet counts from one, so the numbers here are
 * one less than the ones everywhere else.
 *
 * A shape names the cell it belongs to and, separately, the box it is drawn in. A
 * shape whose cell went away is dropped, matching the comment the other part drops
 * by the same rule. A box corner left standing on a row that went closes up against
 * the rows that went, the top landing on the first row still there and the bottom
 * on the last one before them.
 *
 * A shape with no cell of its own, a form control rather than a note, keeps its
 * place in the part and only its box moves.
 */
export async function* shiftVmlAnchors(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  // A shape says which cell it belongs to after it says where its box is, so
  // whether it survives is only known at its closing tag. One shape is small.
  let shape: XmlEvent[] | undefined;

  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      if (event.type === "open" && event.name === Element.Shape) {
        shape = [event];
        continue;
      }

      if (shape === undefined) {
        moved.push(event);
        continue;
      }

      shape.push(event);

      if (event.type === "close" && event.name === Element.Shape) {
        moved.push(...movedShape(shape, shift));
        shape = undefined;
      }
    }

    yield moved;
  }
}

function movedShape(shape: readonly XmlEvent[], shift: RowShift): readonly XmlEvent[] {
  const moved: XmlEvent[] = [];
  let element: string | undefined;

  for (const event of shape) {
    if (event.type === "open") {
      element = event.name;
    } else if (event.type === "close") {
      element = undefined;
    } else if (element === Element.Row) {
      const row = movedRow(Number(event.text) + 1, shift);

      if (row === undefined) {
        return [];
      }

      moved.push({ ...event, text: String(row - 1) });
      continue;
    } else if (element === Element.Anchor) {
      moved.push({ ...event, text: movedAnchor(event.text, shift) });
      continue;
    }

    moved.push(event);
  }

  return moved;
}

function movedAnchor(anchor: string, shift: RowShift): string {
  const corners = anchor.split(",");

  return corners
    .map((corner, index) => {
      const which = ANCHOR_CORNERS.get(index);

      if (which === undefined) {
        return corner;
      }

      return corner.replace(NUMBER, String(cornerRow(Number(corner) + 1, which, shift) - 1));
    })
    .join(",");
}

function cornerRow(row: number, corner: "from" | "to", shift: RowShift): number {
  const moved = movedRow(row, shift);

  if (moved !== undefined) {
    return moved;
  }

  return Math.max(1, corner === "from" ? shift.at : shift.at - 1);
}
