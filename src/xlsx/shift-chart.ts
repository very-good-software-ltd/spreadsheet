import type { XmlEvent } from "../xml/xml-reader";
import { type RowShift, shiftFormula } from "./shift-formula";

const Element = {
  Reference: "c:f",
} as const;

/**
 * A chart part's events with every range it reads moved with the rows.
 *
 * A chart names the sheet it reads in the reference itself, so a series pointing
 * at another sheet is left as it is, wherever the chart is drawn.
 *
 * A reference with nothing left to point at becomes `#REF!`, which is what Excel
 * writes when the rows a series read are deleted.
 *
 * Throws on a reference we cannot move with confidence, the same shape a formula
 * on a sheet refuses.
 */
export async function* shiftChartReferences(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  let reference: string | undefined;

  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      if (event.type === "open" && event.name === Element.Reference) {
        reference = "";
      } else if (event.type === "text" && reference !== undefined) {
        // A reference long enough to straddle a chunk boundary arrives as more than
        // one text event, and half of one moves to nonsense, so it is held whole.
        reference += event.text;
        continue;
      } else if (event.type === "close" && event.name === Element.Reference) {
        moved.push({ type: "text", text: shiftFormula(reference ?? "", shift, shift.sheet) });
        reference = undefined;
      }

      moved.push(event);
    }

    yield moved;
  }
}
