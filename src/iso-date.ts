// An ISO 8601 date string. A date and time with no time zone is read as UTC, so
// it matches the dates we build from serial numbers, which are UTC too. A date
// on its own is already UTC by the JavaScript spec.
export function parseIsoDate(text: string): Date {
  const hasTime = text.includes("T");
  const hasZone = /[Zz]$|[+-]\d\d:?\d\d$/.test(text);
  return new Date(hasTime && !hasZone ? `${text}Z` : text);
}
