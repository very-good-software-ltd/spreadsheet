/** A rectangular area of one sheet. */
export interface Region {
  readonly sheet: string;

  /** One-based, matching `Row.number`. */
  readonly firstRow: number;
  readonly lastRow: number;

  /** Zero-based, matching `Cell.columnIndex`. */
  readonly firstColumnIndex: number;
  readonly lastColumnIndex: number;
}

/** A region the file's author named, so callers can address it by that name. */
export interface NamedRegion extends Region {
  readonly name: string;
}
