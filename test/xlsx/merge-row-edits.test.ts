import { describe, expect, it } from "vitest";
import type { CellInput } from "../../src/cell-input";
import type { RowSource } from "../../src/editor";
import { type CellEdit, mergeRowEdits, type RowBlock } from "../../src/xlsx/merge-row-edits";
import type { RowEdit } from "../../src/xlsx/write-sheet";

function cell(row: number, column: number, value: CellInput, order: number): CellEdit {
  return { row, column, value, order };
}

function block(startRow: number, rows: RowSource, order: number, inheritFrom?: number): RowBlock {
  return { startRow, rows, inheritFrom, order };
}

async function merged(cells: readonly CellEdit[], blocks: readonly RowBlock[]): Promise<unknown[]> {
  const collected: unknown[] = [];

  for await (const edit of mergeRowEdits(cells, blocks)) {
    collected.push(plain(edit));
  }

  return collected;
}

function plain(edit: RowEdit): unknown {
  return {
    number: edit.number,
    cells: Object.fromEntries(edit.cells),
    ...(edit.inheritFrom === undefined ? {} : { inheritFrom: edit.inheritFrom }),
  };
}

describe("mergeRowEdits", () => {
  it("yields nothing for nothing", async () => {
    expect(await merged([], [])).toEqual([]);
  });

  it("groups single cells by row, in ascending order", async () => {
    const cells = [cell(3, 0, "c", 1), cell(1, 1, "b", 2), cell(1, 0, "a", 3)];

    expect(await merged(cells, [])).toEqual([
      { number: 1, cells: { 0: "a", 1: "b" } },
      { number: 3, cells: { 0: "c" } },
    ]);
  });

  it("numbers a block's rows from its start row", async () => {
    expect(await merged([], [block(5, [["a"], ["b"]], 1)])).toEqual([
      { number: 5, cells: { 0: "a" } },
      { number: 6, cells: { 0: "b" } },
    ]);
  });

  it("interleaves blocks and single cells by row number", async () => {
    const cells = [cell(1, 0, "one", 1), cell(9, 0, "nine", 2)];

    expect(await merged(cells, [block(4, [["four"], ["five"]], 3)])).toEqual([
      { number: 1, cells: { 0: "one" } },
      { number: 4, cells: { 0: "four" } },
      { number: 5, cells: { 0: "five" } },
      { number: 9, cells: { 0: "nine" } },
    ]);
  });

  it("combines two sources landing on the same row, column by column", async () => {
    expect(await merged([cell(2, 0, "left", 1)], [block(2, [[undefined, "right"]], 2)])).toEqual([
      { number: 2, cells: { 0: "left", 1: "right" } },
    ]);
  });

  it("lets the later call win where both cover the same cell", async () => {
    expect(await merged([cell(2, 0, "earlier", 1)], [block(2, [["later"]], 2)])).toEqual([
      { number: 2, cells: { 0: "later" } },
    ]);
    expect(await merged([cell(2, 0, "later", 2)], [block(2, [["earlier"]], 1)])).toEqual([
      { number: 2, cells: { 0: "later" } },
    ]);
  });

  it("lets the last of two overlapping blocks win", async () => {
    const blocks = [block(1, [["first"], ["first"]], 1), block(2, [["second"]], 2)];

    expect(await merged([], blocks)).toEqual([
      { number: 1, cells: { 0: "first" } },
      { number: 2, cells: { 0: "second" } },
    ]);
  });

  it("lets the last of two single-cell edits win", async () => {
    expect(await merged([cell(1, 0, "earlier", 1), cell(1, 0, "later", 2)], [])).toEqual([
      { number: 1, cells: { 0: "later" } },
    ]);
  });

  it("skips a gap in a row but keeps an explicit blank", async () => {
    expect(await merged([], [block(1, [[undefined, null, "c"]], 1)])).toEqual([
      { number: 1, cells: { 1: null, 2: "c" } },
    ]);
  });

  it("carries the row a block inherits formatting from", async () => {
    expect(await merged([], [block(7, [["a"]], 1, 6)])).toEqual([{ number: 7, cells: { 0: "a" }, inheritFrom: 6 }]);
  });

  it("reads an async row source", async () => {
    async function* rows(): AsyncIterable<readonly CellInput[]> {
      yield ["a"];
      yield ["b"];
    }

    expect(await merged([], [block(1, rows(), 1)])).toEqual([
      { number: 1, cells: { 0: "a" } },
      { number: 2, cells: { 0: "b" } },
    ]);
  });

  it("reads a row source only one row ahead", async () => {
    let pulled = 0;
    async function* rows(): AsyncIterable<readonly CellInput[]> {
      for (let i = 0; i < 100; i += 1) {
        pulled += 1;
        yield [i];
      }
    }

    const stream = mergeRowEdits([], [block(1, rows(), 1)])[Symbol.asyncIterator]();
    await stream.next();

    expect(pulled).toBeLessThanOrEqual(2);
  });
});
