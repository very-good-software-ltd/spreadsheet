import { parseIsoDate } from "./iso-date";

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
 * A date to write into a cell. Build one with `date`.
 *
 * `value` is the calendar date and time held as UTC, which is the convention the
 * serial numbers in a file follow.
 */
export class SpreadsheetDate {
  constructor(readonly value: Date) {}
}

/**
 * A value to write into a cell.
 *
 * `null` blanks the cell and keeps whatever formatting it already had.
 * There is no error variant: an error is something a formula produces, so write
 * one with `formula("NA()")` rather than as a value.
 * A `Date` is deliberately not accepted, see `date`.
 */
export type CellInput = number | string | boolean | null | Formula | SpreadsheetDate;

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

/**
 * A date for a cell, from an ISO 8601 string such as `"2026-03-01"` or
 * `"2026-03-01T14:30"`, or from its parts.
 *
 * The month is 1 to 12, not JavaScript's 0 to 11.
 *
 * A string with no time zone is read as the calendar date and time it names. One
 * with a zone is read as the instant it names. Throws if the string is not a date,
 * or if the parts do not name a real day, so `date(2026, 2, 30)` is an error
 * rather than the 2nd of March.
 *
 * A `Date` is not accepted, here or as a cell value, and this is deliberate. A
 * cell holds a calendar date with no time zone while a `Date` is an instant, so
 * reading one means picking a zone, and either choice is wrong for some caller.
 * `new Date(2026, 2, 1)` is local midnight, which west of UTC is the last day of
 * February. If you hold a `Date` and want the calendar values it has in UTC, pass
 * `instant.toISOString()`. For its local values, pass its parts.
 */
export function date(isoText: string): SpreadsheetDate;
export function date(
  year: number,
  month: number,
  day: number,
  hours?: number,
  minutes?: number,
  seconds?: number,
): SpreadsheetDate;
export function date(
  first: string | number,
  month?: number,
  day?: number,
  hours = 0,
  minutes = 0,
  seconds = 0,
): SpreadsheetDate {
  if (typeof first === "string") {
    return new SpreadsheetDate(fromText(first));
  }
  if (month === undefined || day === undefined) {
    throw new Error("A date from parts needs a year, a month and a day");
  }
  return new SpreadsheetDate(fromParts(first, month, day, hours, minutes, seconds));
}

function fromText(isoText: string): Date {
  const parsed = parseIsoDate(isoText);

  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Not a date: "${isoText}". Expected an ISO 8601 date like "2026-03-01"`);
  }

  return parsed;
}

function fromParts(year: number, month: number, day: number, hours: number, minutes: number, seconds: number): Date {
  const parts = [year, month, day, hours, minutes, seconds];
  if (parts.some((part) => !Number.isInteger(part))) {
    throw new Error(`Not a date: ${parts.join(", ")}. Every part has to be a whole number`);
  }

  const built = new Date(Date.UTC(year, month - 1, day, hours, minutes, seconds));

  // Date.UTC rolls a part that is out of range into the next unit, so the 30th of
  // February becomes the 2nd of March without complaint. Comparing back is the
  // only way to tell a real day from one that rolled over.
  const rolled =
    built.getUTCFullYear() !== year ||
    built.getUTCMonth() !== month - 1 ||
    built.getUTCDate() !== day ||
    built.getUTCHours() !== hours ||
    built.getUTCMinutes() !== minutes ||
    built.getUTCSeconds() !== seconds;

  if (rolled) {
    throw new Error(`Not a date: ${parts.join(", ")}. The month is 1 to 12, and that day does not exist`);
  }

  return built;
}
