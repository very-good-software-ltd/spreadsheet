import { cellReference, columnIndexOf } from "../cell-reference";
import type { XmlEvent, XmlOpen } from "../xml/xml-reader";
import { movedRow, type RowShift, shiftFormula } from "./shift-formula";
import type { SourceRow } from "./write-sheet";

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

/**
 * Where `row` ends up once the rows have moved, or `undefined` when it was one of
 * the rows taken out.
 *
 * Its cells move with it, and any formula in them is rewritten, since a formula in
 * a row that moved refers to cells that moved too.
 */
export function movedSourceRow(row: SourceRow, shift: RowShift): SourceRow | undefined {
  const number = movedRow(row.number, shift);
  if (number === undefined) {
    return undefined;
  }

  return {
    number,
    attributes: { ...row.attributes, [Attribute.Reference]: String(number) },
    cells: row.cells.map((cell) => ({
      columnIndex: cell.columnIndex,
      attributes: movedAttributes(cell.attributes, cell.attributes[Attribute.Reference], shift),
      inner: movedInner(cell.inner, shift),
    })),
  };
}

/** One event from outside the rows, with any reference on it moved. */
export function movedSheetEvent(event: XmlEvent, shift: RowShift): XmlEvent {
  if (event.type !== "open") {
    return event;
  }

  if (event.name === Element.Extensions) {
    throw new Error("This worksheet holds an extension list, whose contents cannot be moved with confidence");
  }

  return shiftedEvent(event, shift);
}

/**
 * Whether anything here points at row `row` or below it.
 *
 * Asked of what sits above a region, because that goes out before the rows are
 * counted and so before it can be known how far anything moves. Something up there
 * pointing down means the count has to be worked out first.
 */
export function pointsAtOrBelow(events: readonly XmlEvent[], row: number): boolean {
  return events.some((event) => eventPointsAtOrBelow(event, row));
}

function eventPointsAtOrBelow(event: XmlEvent, row: number): boolean {
  const texts =
    event.type === "open"
      ? REFERENCE_ATTRIBUTES.map((name) => event.attributes[name]).filter((value) => value !== undefined)
      : event.type === "text"
        ? [event.text]
        : [];

  return texts.some((text) => mentionsRowAtOrBelow(text, row));
}

// A reference we cannot read is treated as pointing below, so an unfamiliar shape
// makes the writer wait for the count rather than write something stale.
function mentionsRowAtOrBelow(text: string, row: number): boolean {
  for (const match of text.matchAll(/\$?[A-Z]{1,3}\$?([0-9]{1,7})/g)) {
    if (Number(match[1]) >= row) {
      return true;
    }
  }

  return false;
}

function movedAttributes(
  attributes: Readonly<Record<string, string>>,
  reference: string | undefined,
  shift: RowShift,
): Readonly<Record<string, string>> {
  return reference === undefined ? attributes : { ...attributes, [Attribute.Reference]: movedCell(reference, shift) };
}

function movedInner(inner: readonly XmlEvent[], shift: RowShift): readonly XmlEvent[] {
  let inFormula = false;

  return inner.map((event) => {
    if (event.type === "open") {
      inFormula = inFormula || FORMULA_ELEMENTS.has(event.name);
      return shiftedEvent(event, shift);
    }
    if (event.type === "close") {
      inFormula = inFormula && !FORMULA_ELEMENTS.has(event.name);
      return event;
    }

    return inFormula ? { ...event, text: shiftFormula(event.text, shift, shift.sheet) } : event;
  });
}
