// MS-DOS packed 1980-01-01, the earliest date the zip format can express, used
// for every entry we create. A real clock would make the same input produce a
// different file on every run, and nothing reads an entry timestamp out of a
// spreadsheet.
export const DOS_EPOCH_DATE = (1 << 5) | 1;
export const DOS_EPOCH_TIME = 0;
