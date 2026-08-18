import { cellReference, columnIndexOf, rowNumberOf } from "../cell-reference";
import type { XmlEvent, XmlOpen } from "../xml/xml-reader";

const Element = {
  Table: "table",
  AutoFilter: "autoFilter",
} as const;

const Attribute = {
  Extent: "ref",
} as const;

// Both elements carry an extent, and the filter's normally stops one row short of
// a totals row. A table with a totals row is never grown, because that row would
// have to move down, so here the two always end on the same row.
const EXTENDED: ReadonlySet<string> = new Set([Element.Table, Element.AutoFilter]);

/** The table part with its extent, and its filter's, ending at `lastRow`. */
export async function* withTableExtent(events: AsyncIterable<XmlEvent>, lastRow: number): AsyncIterable<XmlEvent> {
  for await (const event of events) {
    yield event.type === "open" && EXTENDED.has(event.name) ? withLastRow(event, lastRow) : event;
  }
}

function withLastRow(event: XmlOpen, lastRow: number): XmlEvent {
  const extent = event.attributes[Attribute.Extent];
  if (extent === undefined) {
    return event;
  }

  const corners = extent.split(":");
  const last = corners[corners.length - 1];
  if (last === undefined) {
    return event;
  }

  const moved = cellReference(lastRow, columnIndexOf(last));
  const first = corners.length > 1 ? corners[0] : cellReference(rowNumberOf(last), columnIndexOf(last));

  return { ...event, attributes: { ...event.attributes, [Attribute.Extent]: `${first}:${moved}` } };
}
