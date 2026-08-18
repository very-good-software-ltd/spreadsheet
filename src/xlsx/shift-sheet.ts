import { cellReference, columnIndexOf } from "../cell-reference";
import type { XmlEvent, XmlOpen } from "../xml/xml-reader";
import { movedRow, type RowShift, shiftFormula } from "./shift-formula";

const Element = {
  Row: "row",
  Cell: "c",
  Break: "brk",
  Extensions: "extLst",
} as const;

// Every element whose text is a formula. A cell's is `f`, a conditional format's
// rule is `formula`, and a data validation has one or two of its own.
const FORMULA_ELEMENTS: ReadonlySet<string> = new Set(["f", "formula", "formula1", "formula2"]);

const Attribute = {
  Reference: "r",
  Extent: "ref",
  Extents: "sqref",
  TopLeftCell: "topLeftCell",
  ActiveCell: "activeCell",
  BreakRow: "id",
} as const;

// Every attribute holding one or more A1 references, wherever it appears. A row's
// `spans` is not one: it counts columns, and no column moves.
const REFERENCE_ATTRIBUTES: readonly string[] = [
  Attribute.Extent,
  Attribute.Extents,
  Attribute.TopLeftCell,
  Attribute.ActiveCell,
];

/**
 * The worksheet's events with every row at or below `shift.at` moved, and every
 * reference on the sheet moved with them.
 *
 * Rows the shift removed are dropped entirely, so a caller downstream sees the
 * sheet as it will be rather than as it was.
 *
 * Throws on anything holding a reference we cannot move with confidence, so a
 * caller can refuse the file rather than write one that is quietly wrong.
 */
export async function* shiftSheetRows(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
): AsyncIterable<readonly XmlEvent[]> {
  let droppingRow = false;
  let inFormula = false;

  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      if (droppingRow) {
        droppingRow = !(event.type === "close" && event.name === Element.Row);
        continue;
      }

      if (event.type === "text") {
        moved.push(inFormula ? { ...event, text: shiftFormula(event.text, shift, shift.sheet) } : event);
        continue;
      }

      if (event.type === "close") {
        inFormula = inFormula && !FORMULA_ELEMENTS.has(event.name);
        moved.push(event);
        continue;
      }

      if (event.name === Element.Extensions) {
        throw new Error("This worksheet holds an extension list, whose contents cannot be moved with confidence");
      }

      if (event.name === Element.Row) {
        const row = movedRow(Number(event.attributes[Attribute.Reference] ?? "0"), shift);
        if (row === undefined) {
          droppingRow = true;
          continue;
        }
        moved.push(withAttribute(event, Attribute.Reference, String(row)));
        continue;
      }

      inFormula = inFormula || FORMULA_ELEMENTS.has(event.name);
      moved.push(shiftedEvent(event, shift));
    }

    if (moved.length > 0) {
      yield moved;
    }
  }
}

function shiftedEvent(event: XmlOpen, shift: RowShift): XmlEvent {
  let moved: XmlEvent = event;

  for (const name of REFERENCE_ATTRIBUTES) {
    const value = event.attributes[name];
    if (value !== undefined) {
      moved = withAttribute(moved, name, shiftFormula(value, shift, shift.sheet));
    }
  }

  if (event.name === Element.Cell) {
    const reference = event.attributes[Attribute.Reference];
    if (reference !== undefined) {
      moved = withAttribute(moved, Attribute.Reference, movedCell(reference, shift));
    }
  }

  // A page break names the row it sits above by number rather than by reference.
  if (event.name === Element.Break) {
    const row = movedRow(Number(event.attributes[Attribute.BreakRow] ?? "0"), shift);
    if (row !== undefined) {
      moved = withAttribute(moved, Attribute.BreakRow, String(row));
    }
  }

  return moved;
}

// A cell in a row that survived is in a row that survived, so it always lands
// somewhere.
function movedCell(reference: string, shift: RowShift): string {
  const row = movedRow(Number(reference.replaceAll(/[$A-Z]/g, "")), shift);

  return row === undefined ? reference : cellReference(row, columnIndexOf(reference));
}

function withAttribute(event: XmlEvent, name: string, value: string): XmlEvent {
  return event.type === "open" ? { ...event, attributes: { ...event.attributes, [name]: value } } : event;
}

/**
 * The events of a worksheet the move did not happen on, with only its formulas
 * rewritten.
 *
 * Its own rows did not move, so its row numbers, its own references and its own
 * ranges are all left exactly as they are. What can be stale is a formula naming
 * the sheet that did move, and only a qualified reference can name it.
 */
export async function* shiftForeignFormulas(
  events: AsyncIterable<readonly XmlEvent[]>,
  shift: RowShift,
  onSheet: string,
): AsyncIterable<readonly XmlEvent[]> {
  let inFormula = false;

  for await (const batch of events) {
    const moved: XmlEvent[] = [];

    for (const event of batch) {
      if (event.type === "open") {
        inFormula = inFormula || FORMULA_ELEMENTS.has(event.name);
      } else if (event.type === "close") {
        inFormula = inFormula && !FORMULA_ELEMENTS.has(event.name);
      } else if (inFormula) {
        moved.push({ ...event, text: shiftFormula(event.text, shift, onSheet) });
        continue;
      }

      moved.push(event);
    }

    yield moved;
  }
}
