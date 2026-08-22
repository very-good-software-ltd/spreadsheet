import type { XmlEvent } from "../xml/xml-reader";
import { type RowShift, shiftFormula } from "./shift-formula";

const Element = {
  CacheDefinition: "pivotCacheDefinition",
  WorksheetSource: "worksheetSource",
  Location: "location",
} as const;

const Attribute = {
  Reference: "ref",
  RefreshOnLoad: "refreshOnLoad",
} as const;

// What a reference with nothing left to point at becomes. A formula can carry it, a
// range attribute cannot, so it is a refusal here rather than a value.
const BROKEN = "#REF!";

/**
 * A pivot cache definition's events with the range it reads moved with the rows.
 *
 * The range is written without a sheet, which the element names separately, so it
 * is read as being on `shift.sheet` and the caller applies this only to a cache
 * that reads that sheet.
 *
 * Throws when the moved rows take the whole source range with them, since a range
 * attribute has no spelling for a reference that points nowhere.
 */
export function shiftPivotSource(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  return withMovedReference(events, Element.WorksheetSource, shift, "the range a pivot table reads");
}

/**
 * A pivot table part's events with the block it is drawn in moved with the rows.
 *
 * Applies to a pivot drawn on the sheet whose rows moved. One reading that sheet
 * from another is not moved by this, and does not move.
 */
export function shiftPivotLocation(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  return withMovedReference(events, Element.Location, shift, "the block a pivot table is drawn in");
}

/**
 * A pivot cache definition's events with the application asked to rebuild it when
 * the file is opened.
 *
 * A cache holds its own copy of the source rows, in a part of its own, and writing
 * to a region leaves that copy describing data that is no longer there. Rebuilding
 * on open is how the copy is made to match again, the same bargain
 * `fullCalcOnLoad` strikes for a stale formula result.
 */
export async function* withCacheRefreshedOnLoad(
  events: AsyncIterable<readonly XmlEvent[]>,
): AsyncIterable<readonly XmlEvent[]> {
  for await (const batch of events) {
    yield batch.map((event) =>
      event.type === "open" && event.name === Element.CacheDefinition
        ? { ...event, attributes: { ...event.attributes, [Attribute.RefreshOnLoad]: "1" } }
        : event,
    );
  }
}

async function* withMovedReference(
  events: AsyncIterable<readonly XmlEvent[]>,
  element: string,
  shift: RowShift,
  what: string,
): AsyncIterable<readonly XmlEvent[]> {
  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      const reference =
        event.type === "open" && event.name === element ? event.attributes[Attribute.Reference] : undefined;

      if (reference === undefined || event.type !== "open") {
        moved.push(event);
        continue;
      }

      const shifted = shiftFormula(reference, shift, shift.sheet);
      if (shifted.includes(BROKEN)) {
        throw new Error(`The rows being taken out of "${shift.sheet}" take all of ${what} with them`);
      }

      moved.push({ ...event, attributes: { ...event.attributes, [Attribute.Reference]: shifted } });
    }

    yield moved;
  }
}
