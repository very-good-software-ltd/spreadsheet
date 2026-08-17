# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [Unreleased]

### Added

- Write `.xlsx` files.
  `workbook.edit()` returns an editor and `editor.save()` streams the result as bytes.
  Reading and writing stay separate: the editor produces a new file and never changes the one you opened.
- Fill a template.
  Every part of the file you do not edit is copied across byte for byte, so charts, pivot tables, drawings, macros and formatting survive.
- Build a file from nothing with `Workbook.create()`.
  It is the same code from the second line on, because from scratch is a template that happens to be empty.
- Write a cell with `worksheet.set("C3", value)`, taking a number, a string, a boolean, a `Date`, `null` to blank it, or `formula(...)`.
  A written cell keeps whatever formatting it already had.
- Write rows with `writeRows(startRow, rows, { inheritFrom })`, or after the last row with `appendRows(rows)`.
  Both take an array, an iterable or an async iterable, and read it only as the output drains, so a generator streams and nothing is held.
- Add a sheet with `editor.addWorksheet(name)`.
- Written files ask the application to recalculate when it opens them, so a total over cells you changed comes out right rather than showing the result the template was saved with.
- Writing a `Date` into a cell that has no date format derives one and keeps the cell's font, fill and border.
- Convert between cell references and coordinates with `cellReference`, `columnIndexOf`, `columnLetters` and `rowNumberOf`.

Worth knowing before you rely on it.
Only `.xlsx` can be written, not `.ods`.
Rows are written over, never inserted, and nothing is pushed down, so a template with a totals block under its data region can have that block overwritten.
`save()` streams rather than returning bytes, so a failure part way through leaves an incomplete file, and calling it twice throws because your row sources have already been read.


### Fixed

- A number format declared inside a differential format (`dxfs`) is no longer read as if it were one of the workbook's own custom formats.
  It could give a format id a date format code it does not have, and make plain numbers using that id read as dates.


## [0.1.0] - 2026-08-05

### Added

- Read Excel `.xlsx` and OpenDocument `.ods` spreadsheets, with `Workbook.open` picking the format from the file.
- Stream a sheet row by row through `worksheet(name).rows()`, so memory stays flat even on very large files.
- Run the same code in Node and the browser, built on web standards with nothing Node specific in the core.
- Open a workbook from a `Uint8Array`, an `ArrayBuffer`, a `ReadableStream` of bytes, or a `Blob`.
  A seekable `Blob`, a browser `File` or a handle from `fs.openAsBlob`, is read in ranges and never held whole.
- Typed cells covering `number`, `string`, `boolean`, `date`, `error`, and `formula`.
  Dates come back as a `Date`, and a formula cell carries its text and its cached result.
- Read a row's cells by column with `row.cell(index)`, each cell carrying a zero-based `columnIndex` and an `"A1"` style `ref`.
- List sheets with `workbook.worksheets`, including whether each one is hidden.
- Ship TypeScript types and an ES module.

[0.1.0]: https://github.com/christophgockel/very-good-spreadsheet/releases/tag/0.1.0
