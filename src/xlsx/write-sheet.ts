import { type CellInput, Formula, SpreadsheetDate } from "../cell-input";
import { cellReference, columnIndexOf } from "../cell-reference";
import { writeXmlEvent, XML_DECLARATION } from "../xml/write-xml";
import type { XmlEvent } from "../xml/xml-reader";
import { dateToSerial } from "./date";
import type { RowShift } from "./shift-formula";
import { insideFormula, movedSheetEvent, movedSourceRow, pointsAtOrBelow } from "./shift-sheet";

const Element = {
  SheetData: "sheetData",
  Dimension: "dimension",
  Row: "row",
  Cell: "c",
  Value: "v",
  Formula: "f",
  InlineString: "is",
  Text: "t",
} as const;

const Attribute = {
  Reference: "r",
  Type: "t",
  Style: "s",
  Spans: "spans",
  Space: "xml:space",
} as const;

const EMPTY_COLUMNS: ReadonlyMap<number, SourceCell> = new Map();

const CellTypeCode = {
  Boolean: "b",
  InlineString: "inlineStr",
} as const;

/** The cells to write into one row, by zero-based column index. */
export type RowCells = ReadonlyMap<number, CellInput>;

export interface RowEdit {
  /** The one-based row number to write into. */
  readonly number: number;

  readonly cells: RowCells;

  /**
   * A row whose formatting the written row copies, for its own row attributes and
   * for cells in columns where the target has no style of its own. Must be at or
   * before `number`, since the sheet is read once from the top.
   */
  readonly inheritFrom?: number;

  /**
   * Whether a missing `inheritFrom` row is acceptable. A caller who named the row
   * gets an error, since they meant a row that is not there. A region infers the
   * row instead of being told it, and a region over rows the sheet never
   * mentions has no formatting to copy and needs none.
   */
  readonly inheritIsOptional?: boolean;
}

export interface RegionWrite {
  /** The rows the region covers in the sheet as it stands. */
  readonly firstRow: number;
  readonly lastRow: number;

  /** Read once, as the region is written, and never held. */
  readonly rows: AsyncIterable<readonly (CellInput | undefined)[]>;

  /** How far the sheet has to move for the region to hold `count` rows. */
  moveFor(count: number): RowShift | undefined;

  /** Told how far the sheet moved, once its rows have been counted. */
  moved(shift: RowShift | undefined): void;
}

export interface SheetWritePlan {
  /** Row edits at known row numbers, in ascending order. */
  readonly positioned: AsyncIterable<RowEdit>;

  /** Rows to place after the last row the sheet already has, in order. */
  readonly appended: AsyncIterable<RowCells>;

  /**
   * Every row number some edit copies formatting from. The sheet is read once, so
   * a row has to be recognised as worth holding while it goes past, which cannot
   * be worked out from the edits alone because they are read lazily.
   */
  readonly inheritedRows: ReadonlySet<number>;

  /**
   * A region to fill, whose rows decide how far the rest of the sheet moves.
   *
   * The rows are read as they are written, so nothing is held, unless something
   * above the region points below it. That has to go out before the rows have been
   * counted, so when it exists the rows are counted first.
   */
  readonly region?: RegionWrite;
}

/** Supplies a style index whose number format renders a date. */
export interface DateStyles {
  /**
   * A style index to put on a date cell. `from` is the cell's existing style
   * index, whose other formatting is kept, and is absent when it has none.
   */
  forDate(from: string | undefined, withTime: boolean): string;
}

export interface SheetWriteContext {
  readonly dateStyles: DateStyles;
  readonly date1904: boolean;
}

export interface SourceCell {
  readonly columnIndex: number;
  readonly attributes: Readonly<Record<string, string>>;
  readonly inner: readonly XmlEvent[];
}

/** A row of the worksheet being rewritten, as it was read from the part. */
export interface SourceRow {
  readonly number: number;
  readonly attributes: Readonly<Record<string, string>>;
  readonly cells: readonly SourceCell[];
}

type SheetPiece =
  | { readonly kind: "prologue"; readonly event: XmlEvent }
  | { readonly kind: "row"; readonly row: SourceRow }
  | { readonly kind: "regionRow"; readonly edit: RowEdit; readonly source: SourceRow | undefined }
  | { readonly kind: "epilogue"; readonly event: XmlEvent };

/**
 * Rewrites a worksheet part, applying `plan` to it.
 *
 * Every event outside the rows passes through untouched, so a template's column
 * widths, merged ranges, conditional formats and panes survive. Inside the rows,
 * only the cells being written are rebuilt, and each keeps its own style.
 */
export async function* writeSheetPart(
  events: AsyncIterable<readonly XmlEvent[]>,
  plan: SheetWritePlan,
  context: SheetWriteContext,
): AsyncIterable<string> {
  yield XML_DECLARATION;

  const positioned = peekable(plan.positioned);
  const inheritable = new Map<number, SourceRow>();
  let lastRowNumber = 0;
  let rowsClosed = false;

  const write = (edit: RowEdit, source: SourceRow | undefined): string => {
    if (edit.inheritFrom === undefined) {
      return writeRow(edit, source, undefined, context);
    }

    const inherited = inheritable.get(edit.inheritFrom);
    if (inherited === undefined) {
      if (edit.inheritIsOptional === true) {
        return writeRow(edit, source, undefined, context);
      }
      throw new Error(
        `Cannot copy formatting from row ${edit.inheritFrom} onto row ${edit.number}: the sheet has no row ${edit.inheritFrom}`,
      );
    }

    return writeRow(edit, source, inherited, context);
  };

  for await (const piece of withRegion(parseSheet(events), plan)) {
    if (piece.kind === "prologue") {
      yield writeXmlEvent(piece.event);
      continue;
    }

    if (piece.kind === "regionRow") {
      for (const edit of await positioned.takeWhile((next) => next.number < piece.edit.number)) {
        yield write(edit, undefined);
      }

      if (piece.source !== undefined && plan.inheritedRows.has(piece.source.number)) {
        inheritable.set(piece.source.number, piece.source);
      }

      yield write(piece.edit, piece.source);
      lastRowNumber = Math.max(lastRowNumber, piece.edit.number);
      continue;
    }

    if (piece.kind === "row") {
      const { row } = piece;

      for (const edit of await positioned.takeWhile((next) => next.number < row.number)) {
        yield write(edit, undefined);
      }

      // Held before the row is written, so a row can be the one it copies from.
      if (plan.inheritedRows.has(row.number)) {
        inheritable.set(row.number, row);
      }

      const next = await positioned.peek();
      if (next !== undefined && next.number === row.number) {
        await positioned.drop();
        yield write(next, row);
      } else {
        yield writeSourceRow(row);
      }

      lastRowNumber = Math.max(lastRowNumber, row.number);
      continue;
    }

    // The rows are done. Anything still pending goes in before the closing tag,
    // and appended rows finally know which numbers they take.
    if (!rowsClosed) {
      rowsClosed = true;

      for (const edit of await positioned.takeWhile(() => true)) {
        yield write(edit, undefined);
        lastRowNumber = Math.max(lastRowNumber, edit.number);
      }

      for await (const cells of plan.appended) {
        lastRowNumber += 1;
        yield write({ number: lastRowNumber, cells }, undefined);
      }
    }

    yield writeXmlEvent(piece.event);
  }
}

// Splits the part at the rows: everything before them, each row buffered whole,
// then everything from the closing tag on. A row is bounded by the column limit,
// so holding one is not holding the sheet.
async function* parseSheet(events: AsyncIterable<readonly XmlEvent[]>): AsyncIterable<SheetPiece> {
  let inRows = false;
  let rowsDone = false;
  let droppingDimension = false;

  let rowNumber = 0;
  let rowAttributes: Readonly<Record<string, string>> = {};
  let cells: SourceCell[] = [];
  let cell: { columnIndex: number; attributes: Readonly<Record<string, string>>; inner: XmlEvent[] } | null = null;
  let cellDepth = 0;

  for await (const batch of events) {
    for (const event of batch) {
      if (droppingDimension) {
        droppingDimension = !(event.type === "close" && event.name === Element.Dimension);
        continue;
      }

      if (cell !== null) {
        if (event.type === "open") {
          cellDepth += 1;
        } else if (event.type === "close") {
          if (cellDepth === 0) {
            cells.push(cell);
            cell = null;
            continue;
          }
          cellDepth -= 1;
        }
        cell.inner.push(event);
        continue;
      }

      if (!inRows) {
        if (!rowsDone && event.type === "open" && event.name === Element.Dimension) {
          droppingDimension = true;
          continue;
        }
        if (event.type === "open" && event.name === Element.SheetData) {
          inRows = true;
          yield { kind: "prologue", event };
          continue;
        }
        yield { kind: rowsDone ? "epilogue" : "prologue", event };
        continue;
      }

      if (event.type === "open" && event.name === Element.Row) {
        rowNumber = Number(event.attributes[Attribute.Reference] ?? "0");
        rowAttributes = event.attributes;
        cells = [];
      } else if (event.type === "open" && event.name === Element.Cell) {
        const ref = event.attributes[Attribute.Reference] ?? "";
        cell = { columnIndex: columnIndexOf(ref), attributes: event.attributes, inner: [] };
        cellDepth = 0;
      } else if (event.type === "close" && event.name === Element.Row) {
        yield { kind: "row", row: { number: rowNumber, attributes: rowAttributes, cells } };
      } else if (event.type === "close" && event.name === Element.SheetData) {
        inRows = false;
        rowsDone = true;
        yield { kind: "epilogue", event };
      }
    }
  }
}

// The row a block copies formatting from is the same row for every row of that
// block, so its columns are indexed once and kept rather than rebuilt a million
// times over.
const columnIndexes = new WeakMap<readonly SourceCell[], ReadonlyMap<number, SourceCell>>();

function byColumnCached(cells: readonly SourceCell[]): ReadonlyMap<number, SourceCell> {
  const held = columnIndexes.get(cells);
  if (held !== undefined) {
    return held;
  }

  const indexed = byColumn(cells);
  columnIndexes.set(cells, indexed);

  return indexed;
}

function byColumn(cells: readonly SourceCell[]): ReadonlyMap<number, SourceCell> {
  return new Map(cells.map((cell) => [cell.columnIndex, cell]));
}

function writeSourceRow(row: SourceRow): string {
  let out = writeXmlEvent({ type: "open", name: Element.Row, attributes: row.attributes });

  for (const cell of row.cells) {
    out += writeSourceCell(cell);
  }

  return out + writeXmlEvent({ type: "close", name: Element.Row });
}

function writeSourceCell(cell: SourceCell): string {
  let out = writeXmlEvent({ type: "open", name: Element.Cell, attributes: cell.attributes });

  for (const event of cell.inner) {
    out += writeXmlEvent(event);
  }

  return out + writeXmlEvent({ type: "close", name: Element.Cell });
}

function writeRow(
  edit: RowEdit,
  source: SourceRow | undefined,
  inherited: SourceRow | undefined,
  context: SheetWriteContext,
): string {
  const existing = byColumn(source?.cells ?? []);
  const inheritedCells = inherited === undefined ? EMPTY_COLUMNS : byColumnCached(inherited.cells);
  const columns = [...new Set([...existing.keys(), ...edit.cells.keys()])].sort((a, b) => a - b);

  // A row the sheet does not have takes the shape of the row it inherits from,
  // so a written row keeps the template's height and row-level format too, not
  // only its cell styles.
  const attributes: Record<string, string> = {
    ...(source?.attributes ?? inherited?.attributes ?? {}),
    [Attribute.Reference]: String(edit.number),
  };
  const first = columns[0];
  const last = columns.at(-1);
  if (Attribute.Spans in attributes && first !== undefined && last !== undefined) {
    attributes[Attribute.Spans] = `${first + 1}:${last + 1}`;
  }

  let out = writeXmlEvent({ type: "open", name: Element.Row, attributes });

  for (const columnIndex of columns) {
    const value = edit.cells.get(columnIndex);
    const cell = existing.get(columnIndex);

    if (value === undefined && cell !== undefined) {
      out += writeSourceCell(cell);
      continue;
    }

    const style = cell?.attributes[Attribute.Style] ?? inheritedCells.get(columnIndex)?.attributes[Attribute.Style];
    out += writeCell(cellReference(edit.number, columnIndex), value ?? null, style, context);
  }

  return out + writeXmlEvent({ type: "close", name: Element.Row });
}

function writeCell(ref: string, value: CellInput, style: string | undefined, context: SheetWriteContext): string {
  const attributes: Record<string, string> = { [Attribute.Reference]: ref };

  if (value instanceof SpreadsheetDate) {
    const serial = dateToSerial(value.value, context.date1904);
    const dateStyle = context.dateStyles.forDate(style, !Number.isInteger(serial));
    return cell({ ...attributes, [Attribute.Style]: dateStyle }, element(Element.Value, String(serial)));
  }

  if (style !== undefined) {
    attributes[Attribute.Style] = style;
  }

  if (value === null) {
    return cell(attributes, "");
  }
  if (value instanceof Formula) {
    return cell(attributes, element(Element.Formula, value.text));
  }
  if (typeof value === "boolean") {
    return cell({ ...attributes, [Attribute.Type]: CellTypeCode.Boolean }, element(Element.Value, value ? "1" : "0"));
  }
  if (typeof value === "number") {
    return cell(attributes, element(Element.Value, numberText(ref, value)));
  }
  if (typeof value !== "string") {
    throw new Error(`Cannot write ${describe(value)} to ${ref}`);
  }

  return cell({ ...attributes, [Attribute.Type]: CellTypeCode.InlineString }, inlineString(ref, value));
}

// Only reachable from untyped code, since CellInput excludes both of these. A
// Date is worth naming, because it is the obvious thing to reach for and the
// reason it is refused is not obvious.
function describe(value: unknown): string {
  if (value instanceof Date) {
    return 'a Date: a cell holds a calendar date with no time zone, so build one with date("2026-03-01") or date(2026, 3, 1)';
  }
  return `a ${typeof value}`;
}

function numberText(ref: string, value: number): string {
  if (Number.isNaN(value)) {
    throw new Error(`Cannot write NaN to ${ref}: a spreadsheet has no such number`);
  }
  if (!Number.isFinite(value)) {
    throw new Error(`Cannot write ${value} to ${ref}: a cell holds only a finite number`);
  }
  return String(value);
}

function cell(attributes: Record<string, string>, inner: string): string {
  return (
    writeXmlEvent({ type: "open", name: Element.Cell, attributes }) +
    inner +
    writeXmlEvent({ type: "close", name: Element.Cell })
  );
}

function element(name: string, text: string): string {
  return (
    writeXmlEvent({ type: "open", name, attributes: {} }) +
    writeXmlEvent({ type: "text", text }) +
    writeXmlEvent({ type: "close", name })
  );
}

const TAB = 0x09;
const LINE_FEED = 0x0a;
const CARRIAGE_RETURN = 0x0d;
const FIRST_PRINTABLE = 0x20;

// XML 1.0 has no way to carry a C0 control character, escaped or not, so text
// holding one cannot be written at all. Tab, line feed and carriage return are
// the three it does allow. Written as a scan rather than a pattern because a
// pattern matching control characters is indistinguishable, to a linter, from
// one that contains them by accident.
function forbiddenXmlCharacter(text: string): number | undefined {
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    if (code < FIRST_PRINTABLE && code !== TAB && code !== LINE_FEED && code !== CARRIAGE_RETURN) {
      return code;
    }
  }
  return undefined;
}

function inlineString(ref: string, text: string): string {
  const forbidden = forbiddenXmlCharacter(text);
  if (forbidden !== undefined) {
    const hex = forbidden.toString(16).toUpperCase().padStart(4, "0");
    throw new Error(`Cannot be written to ${ref}: the text holds U+${hex}, which XML has no way to represent`);
  }

  // Without this marker a reader is free to trim the surrounding whitespace, and
  // Excel writes it for the same reason.
  const attributes = text.trim() === text ? {} : { [Attribute.Space]: "preserve" };

  return (
    writeXmlEvent({ type: "open", name: Element.InlineString, attributes: {} }) +
    writeXmlEvent({ type: "open", name: Element.Text, attributes }) +
    writeXmlEvent({ type: "text", text }) +
    writeXmlEvent({ type: "close", name: Element.Text }) +
    writeXmlEvent({ type: "close", name: Element.InlineString })
  );
}

interface Peekable<T> {
  peek(): Promise<T | undefined>;
  drop(): Promise<void>;
  takeWhile(predicate: (value: T) => boolean): Promise<readonly T[]>;
}

function peekable<T>(source: AsyncIterable<T>): Peekable<T> {
  const iterator = source[Symbol.asyncIterator]();
  let head: T | undefined;
  let loaded = false;

  const load = async (): Promise<T | undefined> => {
    if (!loaded) {
      const { done, value } = await iterator.next();
      head = done ? undefined : value;
      loaded = true;
    }
    return head;
  };

  const forget = (): void => {
    loaded = false;
    head = undefined;
  };

  return {
    peek: load,
    async drop(): Promise<void> {
      await load();
      forget();
    },
    async takeWhile(predicate): Promise<readonly T[]> {
      const taken: T[] = [];
      for (;;) {
        const value = await load();
        if (value === undefined || !predicate(value)) {
          return taken;
        }
        taken.push(value);
        forget();
      }
    },
  };
}

// Fills the plan's region as the sheet goes past it, and moves everything after it
// by however many rows turned up.
//
// The rows above the region are written before the ones inside it, so they go out
// before the count is known. When none of them points below the region that is
// fine and nothing is held. When one does, the rows are counted first, which is
// the only case where a region costs memory in proportion to its data.
async function* withRegion(pieces: AsyncIterable<SheetPiece>, plan: SheetWritePlan): AsyncIterable<SheetPiece> {
  const region = plan.region;
  if (region === undefined) {
    yield* pieces;
    return;
  }

  const source = pieces[Symbol.asyncIterator]();
  const above: SheetPiece[] = [];
  let overshot: SheetPiece | undefined;

  for (;;) {
    const next = await source.next();
    if (next.done === true) {
      break;
    }
    if (next.value.kind === "row" && next.value.row.number >= region.firstRow) {
      overshot = next.value;
      break;
    }
    above.push(next.value);
  }

  const held = pointsAtOrBelow(above.flatMap(eventsOf), region.firstRow) ? await drainRows(region.rows) : undefined;
  let shift = held === undefined ? undefined : region.moveFor(held.length);

  if (shift !== undefined) {
    region.moved(shift);
  }

  // A conditional format's rule and a data validation's bounds hold their formula
  // as text, so what has been seen carries from one event to the next.
  let inFormula = false;
  const move = (piece: SheetPiece): SheetPiece | undefined => {
    if (shift === undefined) {
      return piece;
    }

    const moved = movedPiece(piece, shift, inFormula);
    if (piece.kind !== "row" && piece.kind !== "regionRow") {
      inFormula = insideFormula(piece.event, inFormula);
    }

    return moved;
  };

  for (const piece of above) {
    const moved = move(piece);
    if (moved !== undefined) {
      yield moved;
    }
  }

  // The rows the region already covers, which the written ones take their cells and
  // formatting from. There are as many as the author drew, so holding them is not
  // holding the sheet.
  const covered = new Map<number, SourceRow>();
  while (overshot?.kind === "row" && overshot.row.number <= region.lastRow) {
    covered.set(overshot.row.number, overshot.row);
    const next = await source.next();
    overshot = next.done === true ? undefined : next.value;
  }

  const height = region.lastRow - region.firstRow + 1;
  const rows = peekable(held === undefined ? region.rows : oneByOne(held));
  let index = 0;

  for (;;) {
    const values = await rows.peek();
    if (values === undefined) {
      break;
    }

    // Read before the next row is asked for. A region hands the same row out again
    // and again with new values in it, so looking ahead overwrites what is here.
    const cells = cellsOf(values);
    await rows.drop();

    const last = (await rows.peek()) === undefined;
    const from = last ? Math.min(height - 1, index) : index < height - 1 ? index : undefined;

    yield {
      kind: "regionRow",
      edit: {
        number: region.firstRow + index,
        cells,
        inheritFrom: region.firstRow,
        inheritIsOptional: true,
      },
      source: from === undefined ? undefined : covered.get(region.firstRow + from),
    };
    index += 1;
  }

  if (shift === undefined) {
    shift = region.moveFor(index);
    region.moved(shift);
  }

  if (overshot !== undefined) {
    const moved = move(overshot);
    if (moved !== undefined) {
      yield moved;
    }
  }

  for (;;) {
    const next = await source.next();
    if (next.done === true) {
      return;
    }

    const moved = move(next.value);
    if (moved !== undefined) {
      yield moved;
    }
  }
}

function movedPiece(piece: SheetPiece, shift: RowShift, inFormula: boolean): SheetPiece | undefined {
  if (piece.kind === "row") {
    const row = movedSourceRow(piece.row, shift);
    return row === undefined ? undefined : { kind: "row", row };
  }
  if (piece.kind === "regionRow") {
    return piece;
  }

  return { kind: piece.kind, event: movedSheetEvent(piece.event, shift, inFormula) };
}

function eventsOf(piece: SheetPiece): XmlEvent[] {
  if (piece.kind === "row") {
    return piece.row.cells.flatMap((cell) => [
      { type: "open", name: Element.Cell, attributes: cell.attributes },
      ...cell.inner,
    ]);
  }

  return piece.kind === "regionRow" ? [] : [piece.event];
}

function cellsOf(values: readonly (CellInput | undefined)[]): RowCells {
  const cells = new Map<number, CellInput>();

  for (const [column, value] of values.entries()) {
    if (value !== undefined) {
      cells.set(column, value);
    }
  }

  return cells;
}

// A region hands the same array out for every row, refilled each time, so holding
// the rows means copying them. Holding the references would leave every row reading
// as the last one written.
async function drainRows(
  source: AsyncIterable<readonly (CellInput | undefined)[]>,
): Promise<(readonly (CellInput | undefined)[])[]> {
  const held: (readonly (CellInput | undefined)[])[] = [];

  for await (const values of source) {
    held.push([...values]);
  }

  return held;
}

async function* oneByOne<T>(items: readonly T[]): AsyncIterable<T> {
  yield* items;
}
