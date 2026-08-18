import type { CellInput } from "../cell-input";
import type { RowSource } from "../editor";
import type { RowEdit } from "./write-sheet";

/**
 * A value together with the position of the call that supplied it, so that where
 * two calls cover the same cell the later one can win.
 */
interface Claim {
  readonly value: CellInput;
  readonly order: number;
}

export interface CellEdit {
  readonly row: number;
  readonly column: number;
  readonly value: CellInput;
  readonly order: number;
}

export interface RowBlock {
  readonly startRow: number;
  readonly rows: RowSource;
  readonly inheritFrom: number | undefined;
  readonly order: number;

  /** See `RowEdit.inheritIsOptional`. */
  readonly inheritIsOptional?: boolean;
}

interface PendingRow {
  readonly number: number;
  readonly cells: ReadonlyMap<number, Claim>;
  readonly inheritFrom: number | undefined;
  readonly inheritIsOptional?: boolean;
}

/**
 * Merges single-cell edits and blocks of rows into one stream of row edits in
 * ascending row order. Where more than one call covers the same cell the last one
 * made wins.
 *
 * Row sources are read one row ahead and no further, so a block of five million
 * rows costs one row of memory here.
 */
export async function* mergeRowEdits(cells: readonly CellEdit[], blocks: readonly RowBlock[]): AsyncIterable<RowEdit> {
  const sources = [lookahead(cellRows(cells)), ...blocks.map((block) => lookahead(blockRows(block)))];

  for (;;) {
    // Each source is carried alongside the row it is offering, so advancing one
    // needs no lookup back into the list by position.
    const waiting = await Promise.all(sources.map(async (source) => ({ source, head: await source.peek() })));
    const next = smallestRowNumber(waiting);

    if (next === undefined) {
      return;
    }

    const claims = new Map<number, Claim>();
    let inheritFrom: number | undefined;
    let inheritIsOptional = false;

    for (const { source, head } of waiting) {
      if (head === undefined || head.number !== next) {
        continue;
      }

      for (const [column, claim] of head.cells) {
        const held = claims.get(column);
        if (held === undefined || claim.order >= held.order) {
          claims.set(column, claim);
        }
      }

      if (head.inheritFrom !== undefined) {
        inheritFrom = head.inheritFrom;
        inheritIsOptional = head.inheritIsOptional === true;
      }
      await source.drop();
    }

    yield {
      number: next,
      cells: new Map([...claims].map(([column, claim]) => [column, claim.value])),
      ...(inheritFrom === undefined ? {} : { inheritFrom, inheritIsOptional }),
    };
  }
}

function smallestRowNumber(waiting: readonly { readonly head: PendingRow | undefined }[]): number | undefined {
  let smallest: number | undefined;

  for (const { head } of waiting) {
    if (head !== undefined && (smallest === undefined || head.number < smallest)) {
      smallest = head.number;
    }
  }

  return smallest;
}

async function* cellRows(cells: readonly CellEdit[]): AsyncIterable<PendingRow> {
  const byRow = new Map<number, Map<number, Claim>>();

  for (const edit of cells) {
    const row = byRow.get(edit.row) ?? new Map<number, Claim>();
    const held = row.get(edit.column);

    if (held === undefined || edit.order >= held.order) {
      row.set(edit.column, { value: edit.value, order: edit.order });
    }
    byRow.set(edit.row, row);
  }

  for (const [number, cells] of [...byRow].sort(([left], [right]) => left - right)) {
    yield { number, cells, inheritFrom: undefined };
  }
}

async function* blockRows(block: RowBlock): AsyncIterable<PendingRow> {
  let number = block.startRow;

  for await (const values of block.rows) {
    yield {
      number,
      cells: claimsOf(values, block.order),
      inheritFrom: block.inheritFrom,
      ...(block.inheritIsOptional === undefined ? {} : { inheritIsOptional: block.inheritIsOptional }),
    };
    number += 1;
  }
}

// A gap in the array, an index holding undefined, leaves that column alone. Only
// an explicit null blanks it.
function claimsOf(values: readonly (CellInput | undefined)[], order: number): ReadonlyMap<number, Claim> {
  const claims = new Map<number, Claim>();

  for (const [column, value] of values.entries()) {
    if (value !== undefined) {
      claims.set(column, { value, order });
    }
  }

  return claims;
}

interface Lookahead<T> {
  peek(): Promise<T | undefined>;
  drop(): Promise<void>;
}

function lookahead<T>(source: AsyncIterable<T>): Lookahead<T> {
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

  return {
    peek: load,
    async drop(): Promise<void> {
      await load();
      loaded = false;
      head = undefined;
    },
  };
}
