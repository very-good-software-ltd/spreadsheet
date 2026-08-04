const DAY_MS = 86_400_000;

// Excel stores a date as a serial number: the count of days since an epoch.
// The workbook picks one of two epochs, flagged by date1904 in workbook.xml.
//
// The 1904 system (originally Mac Excel, whose clock counted from 1904) is a plain offset from 1904-01-01.
//
// The 1900 system (the Windows default) carries a deliberate bug:
// Excel treats 1900 as a leap year, so serial 60 is a 1900-02-29 that never existed.
// Microsoft kept the error for bug-for-bug compatibility with Lotus 1-2-3.
// Every date from 1900-03-01 on is therefore shifted a day late.
// Counting from 1899-12-30 cancels that shift for those dates.
// Serials below 60 predate the phantom day and never got the extra day, so they count from 1899-12-31 instead.
export function serialToDate(serial: number, date1904: boolean): Date {
  if (date1904) {
    return new Date(Date.UTC(1904, 0, 1) + serial * DAY_MS);
  }

  const base = serial < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);

  return new Date(base + serial * DAY_MS);
}
