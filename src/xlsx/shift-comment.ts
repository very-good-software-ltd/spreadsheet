import { cellReference, columnIndexOf, rowNumberOf } from "../cell-reference";
import type { XmlEvent } from "../xml/xml-reader";
import { movedRow, type RowShift } from "./shift-formula";

const Element = {
  Comment: "comment",
} as const;

const Attribute = {
  Reference: "ref",
} as const;

/**
 * A comments part's events with every comment at or below the moved rows moved
 * with them.
 *
 * A comment whose cell went away goes with it, text and all, which is what Excel
 * does when a row carrying one is deleted. The shape that draws its box lives in a
 * separate VML part and is dropped there by the same rule.
 */
export async function* shiftCommentRefs(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  let dropping = 0;

  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      if (dropping > 0) {
        if (event.type === "open") {
          dropping += 1;
        } else if (event.type === "close") {
          dropping -= 1;
        }
        continue;
      }

      if (event.type === "open" && event.name === Element.Comment) {
        const reference = event.attributes[Attribute.Reference] ?? "";
        const row = movedRow(rowNumberOf(reference), shift);

        if (row === undefined) {
          dropping = 1;
          continue;
        }

        const attributes = { ...event.attributes, [Attribute.Reference]: cellReference(row, columnIndexOf(reference)) };
        moved.push({ ...event, attributes });
        continue;
      }

      moved.push(event);
    }

    yield moved;
  }
}
