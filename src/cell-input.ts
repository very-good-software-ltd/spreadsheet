/**
 * A formula to write into a cell, as its text. Build one with `formula`.
 *
 * A class rather than a shape so a plain string can never be mistaken for one,
 * and so the text passes through a single place where a format that spells
 * formulas differently could translate it.
 */
export class Formula {
  constructor(readonly text: string) {}
}

/**
 * A value to write into a cell.
 *
 * `null` blanks the cell and keeps whatever formatting it already had.
 * A `Date` is written as a serial number, with a date format applied.
 * There is no error variant: an error is something a formula produces, so write
 * one with `formula("NA()")` rather than as a value.
 */
export type CellInput = number | string | boolean | Date | null | Formula;

/**
 * A formula from its text, in the A1 form a spreadsheet shows, for example
 * `SUM(A1:B1)`. A leading `=` is accepted and stripped, since a file stores the
 * text without one.
 *
 * The text is written through unparsed and unchecked, so a mistake in it surfaces
 * as the spreadsheet application offering to repair the file, not as an error
 * here. Throws only if there is no formula left after trimming.
 */
export function formula(text: string): Formula {
  const trimmed = text.trim();
  const body = (trimmed.startsWith("=") ? trimmed.slice(1) : trimmed).trim();

  if (body === "") {
    throw new Error("A formula cannot be empty");
  }

  return new Formula(body);
}
