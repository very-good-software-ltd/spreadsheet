import type { CellInput } from "./cell-input";

/**
 * Rows to write, each an array of values by column.
 *
 * Nothing is read from it until the file is saved, so a generator streams and an
 * array simply works. A gap in a row, an index holding `undefined`, leaves that
 * column as it was, where `null` blanks it.
 */
export type RowSource =
  | Iterable<readonly (CellInput | undefined)[]>
  | AsyncIterable<readonly (CellInput | undefined)[]>;

export interface WriteRowsOptions {
  /**
   * A row whose formatting the written rows copy, for its height and row format
   * and for cells in columns where the target has no style of its own. Must be at
   * or before `startRow`, since the sheet is read once from the top.
   */
  readonly inheritFrom?: number;
}

export interface WorksheetEditor {
  /**
   * Write a value into the cell at `ref`, for example `"C3"`. The cell keeps
   * whatever formatting it already had. Throws straight away if `ref` is not a
   * cell reference.
   */
  set(ref: string, value: CellInput): this;

  /**
   * Write `rows` starting at the one-based `startRow`, overwriting the cells they
   * cover and leaving every other cell in those rows as it was.
   *
   * Rows are not inserted and nothing is pushed down, so a row already at that
   * position is written over.
   */
  writeRows(startRow: number, rows: RowSource, options?: WriteRowsOptions): this;

  /** Write `rows` after the last row the sheet already has. */
  appendRows(rows: RowSource): this;
}

export interface Editor {
  /** The editor for one worksheet, by name or by zero-based position. */
  worksheet(nameOrIndex: string | number): WorksheetEditor;

  /**
   * The edited workbook as a stream of bytes. Row sources are read as it drains,
   * so nothing is held.
   *
   * A failure part way through, including one thrown by a caller's own row
   * source, errors the stream and leaves a truncated archive that no reader will
   * open. Throws if called more than once, since an already-read row source would
   * silently produce a valid file with rows missing.
   */
  save(): ReadableStream<Uint8Array>;
}
