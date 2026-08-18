export interface RowShift {
  /** The sheet whose rows moved. An unqualified reference means the sheet it is written on. */
  readonly sheet: string;

  /** The first row of the block that appeared or went away. */
  readonly at: number;

  /** How many rows appeared, negative for rows that went away. */
  readonly by: number;
}

// Excel writes what a reference cannot reach any more into the formula itself.
const BROKEN = "#REF!";

// A sheet name is quoted when it holds a space or punctuation, and a quote inside
// it is written twice.
const SHEET = String.raw`(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)`;
const CELL = String.raw`\$?[A-Z]{1,3}\$?[0-9]{1,7}`;
const REFERENCE = new RegExp(String.raw`(${SHEET}!)?(${CELL})(?::(${CELL}))?`, "y");

// Two shapes we will not move. A reference spanning sheets carries rows for all of
// them, and a whole row reference has no column to tell it apart from a number.
const SHEET_RANGE = new RegExp(String.raw`${SHEET}:${SHEET}!`, "y");
const WHOLE_ROWS = /\$?[0-9]{1,7}:\$?[0-9]{1,7}/y;

// A reference never runs on from a name, and is never the name of a function, so
// what sits either side of it decides whether it is one at all.
const BEFORE = /[A-Za-z0-9_.$!\]]/;
const AFTER = /[A-Za-z0-9_.(]/;

/**
 * `text` with every reference to `shift.sheet` moved, in the A1 spelling xlsx uses
 * and without a leading `=`.
 *
 * A reference with nothing left to point at becomes `#REF!`, which is what Excel
 * itself writes. Text in quotes, structured references, defined names, function
 * names and references into other workbooks are left as they are.
 *
 * Throws on a reference we cannot move with confidence, so a caller can refuse the
 * file rather than write one that is quietly wrong.
 */
export function shiftFormula(text: string, shift: RowShift, onSheet: string): string {
  let result = "";
  let position = 0;

  while (position < text.length) {
    const character = text[position] ?? "";

    if (character === '"') {
      const end = endOfQuoted(text, position);
      result += text.slice(position, end);
      position = end;
      continue;
    }

    if (character === "[") {
      const end = text.indexOf("]", position);
      const stop = end < 0 ? text.length : end + 1;
      result += text.slice(position, stop);
      position = stop;
      continue;
    }

    if (isBoundary(text[position - 1])) {
      refuseUnmovable(text, position);

      REFERENCE.lastIndex = position;
      const match = REFERENCE.exec(text);
      if (match !== null && !AFTER.test(text[REFERENCE.lastIndex] ?? "")) {
        result += moved(match, shift, onSheet);
        position = REFERENCE.lastIndex;
        continue;
      }
    }

    result += character;
    position += 1;
  }

  return result;
}

function refuseUnmovable(text: string, position: number): void {
  SHEET_RANGE.lastIndex = position;
  if (SHEET_RANGE.test(text)) {
    throw new Error(`"${text}" refers to a range of sheets, which cannot be moved with confidence`);
  }

  WHOLE_ROWS.lastIndex = position;
  if (WHOLE_ROWS.test(text) && !AFTER.test(text[WHOLE_ROWS.lastIndex] ?? "")) {
    throw new Error(`"${text}" refers to whole rows, which cannot be moved with confidence`);
  }
}

function moved(match: RegExpExecArray, shift: RowShift, onSheet: string): string {
  const [whole, qualifier, first, last] = match;
  const sheet = qualifier === undefined ? onSheet : unquote(qualifier.slice(0, -1));

  if (sheet !== shift.sheet || first === undefined) {
    return whole;
  }

  const prefix = qualifier ?? "";
  if (last === undefined) {
    const row = movedRow(rowOf(first), shift);
    return row === undefined ? BROKEN : prefix + withRow(first, row);
  }

  const range = movedRange(rowOf(first), rowOf(last), shift);
  return range === undefined ? BROKEN : `${prefix}${withRow(first, range[0])}:${withRow(last, range[1])}`;
}

/**
 * Where `row` ends up, or `undefined` when it went away with the rows that were
 * removed and has nowhere left to be.
 */
export function movedRow(row: number, shift: RowShift): number | undefined {
  if (shift.by > 0) {
    return row >= shift.at ? row + shift.by : row;
  }

  const lastGone = shift.at - shift.by - 1;
  if (row > lastGone) {
    return row + shift.by;
  }

  return row < shift.at ? row : undefined;
}

// Only a range with nothing left at either end breaks. One with a single end inside
// the rows that went away closes up against them, which is what Excel does.
function movedRange(first: number, last: number, shift: RowShift): [number, number] | undefined {
  if (shift.by > 0) {
    return [first >= shift.at ? first + shift.by : first, last >= shift.at ? last + shift.by : last];
  }

  const lastGone = shift.at - shift.by - 1;
  const startGone = first >= shift.at && first <= lastGone;
  const endGone = last >= shift.at && last <= lastGone;

  if (startGone && endGone) {
    return undefined;
  }

  const start = startGone ? shift.at : survivor(first, shift);
  const end = endGone ? shift.at - 1 : survivor(last, shift);

  return [start, end];
}

function survivor(row: number, shift: RowShift): number {
  return row < shift.at ? row : row + shift.by;
}

function endOfQuoted(text: string, start: number): number {
  for (let position = start + 1; position < text.length; position += 1) {
    if (text[position] !== '"') {
      continue;
    }
    if (text[position + 1] === '"') {
      position += 1;
      continue;
    }
    return position + 1;
  }

  return text.length;
}

function isBoundary(character: string | undefined): boolean {
  return character === undefined || !BEFORE.test(character);
}

function unquote(sheet: string): string {
  return sheet.startsWith("'") ? sheet.slice(1, -1).replaceAll("''", "'") : sheet;
}

function rowOf(cell: string): number {
  return Number(cell.replaceAll(/[$A-Z]/g, ""));
}

function withRow(cell: string, row: number): string {
  return cell.replace(/[0-9]+$/, String(row));
}
