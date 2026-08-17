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

// The inverse. A serial carries no timezone and a Date is an instant, so the
// conversion has to pick one, and serialToDate builds from UTC. Reading the
// instant back in UTC is what keeps the two consistent. A caller who builds a
// Date from local parts is west of UTC will therefore write the previous day.
//
// In the 1900 system the phantom leap day makes the mapping many-to-one: serials
// 59 and 60 both read as 1900-02-28, since serial 60 is a 1900-02-29 that never
// existed. Writing that date picks 59, the serial that means it.
// A date from 1900-03-01 on is 61 or more days past 1899-12-30, which is where
// the shifted count and the true count agree again.
export function dateToSerial(date: Date, date1904: boolean): number {
  const time = date.getTime();

  if (date1904) {
    return (time - Date.UTC(1904, 0, 1)) / DAY_MS;
  }

  const shifted = (time - Date.UTC(1899, 11, 30)) / DAY_MS;

  return shifted >= 61 ? shifted : (time - Date.UTC(1899, 11, 31)) / DAY_MS;
}
