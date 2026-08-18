# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [Unreleased]

### Changed

- Filling a region allocates one array per row instead of three, which lowered its peak memory by about an eighth on a million rows. Nothing about the output changes.

### Added

- A `region` mode in the write benchmark, so the memory cost of filling a region sits next to the streaming and load-everything paths rather than being described in prose. `npm run benchmark -- --write=1000000`.

## [0.3.0] - 2026-08-18

### Added

- Fill a template by the names its author gave it, with `worksheet.writeRegion("Lines", rows)`.
  Select a range in Excel, name it in the Name Box, and write to that name. The template can then move under your code without breaking it, which a cell reference cannot survive.
  The region ends up exactly as tall as the data you give it, and the sheet moves around it. More rows pushes everything below down, fewer pulls it up. This is the model every other templating tool has, and the same thing Insert and Delete do in Excel.
  Everything pointing at the moved rows moves with them: formulas on that sheet and on every other sheet, merged ranges, conditional formats, data validations, hyperlinks, filters, page breaks, frozen panes, table extents and defined names. A total written `=SUM(C9:C11)` over a three-row region becomes `=SUM(C9:C13)` when five rows arrive, and `=SUM(C9:C9)` when one does.
  No rows at all leaves one empty formatted row, since a region with nothing in it would take every reference to it down with it.
  Charts and images move with the rows. A shape anchored below the region comes down or up with it, and one anchored across it stretches, matching what Excel does when you insert rows by hand.
  Where something on the sheet cannot be moved, `save` throws and names it rather than writing a file that is quietly wrong: a comment below the region, a pivot table reading from it, a formula spanning a range of sheets or naming whole rows, a sheet carrying an extension list, or a shape standing only on rows that are going away.
  A worksheet sees the names it owns before the workbook-wide ones, matching Excel. `editor.writeRegion` sees only the workbook-wide ones and writes into whichever sheet the name points at.
  A name that cannot be written into says what it is, whether a formula, a print area, a whole column or a reference Excel broke when a sheet was deleted, rather than reporting that the name is missing.
  Two limits: one region per worksheet per save, and the rows given to a region are read in full before the file is written, so a region is not the place for a million rows. `writeRows` and `appendRows` are unchanged, move nothing, and stay flat in memory.
  Matching a header row by its text is not something we will do.
- Address an Excel Table by its name, the same way as a named region.
  `writeRegion("Sales", rows)` writes the rows between the table's header and its totals row, since neither is a place for your data.
  The table grows or shrinks to fit, extending its own range and its filter, so a total written as `=SUM(Sales[Amount])` keeps covering everything. A totals row is no obstacle: it moves down with everything else.

## [0.2.0] - 2026-08-17

### Added

- Write `.xlsx` files.
  `workbook.edit()` returns an editor and `editor.save()` streams the result as bytes.
  Reading and writing stay separate: the editor produces a new file and never changes the one you opened.
- Fill a template.
  Every part of the file you do not edit is copied across byte for byte, so charts, pivot tables, drawings, macros and formatting survive.
- Build a file from nothing with `Workbook.create()`.
  It is the same code from the second line on, because from scratch is a template that happens to be empty.
- Write a cell with `worksheet.set("C3", value)`, taking a number, a string, a boolean, `null` to blank it, `formula(...)` or `date(...)`.
  A written cell keeps whatever formatting it already had.
- Build a date with `date("2026-03-01")` or `date(2026, 3, 1, 14, 30)`, with a month from 1 to 12.
  A `Date` is not accepted as a cell value on purpose. A cell holds a calendar date with no time zone while a `Date` is an instant, so `new Date(2026, 2, 1)` is local midnight and west of UTC is the last day of February.
  `date` also refuses a day that does not exist rather than rolling it over, so `date(2026, 2, 30)` is an error and not the 2nd of March.
- Write rows with `writeRows(startRow, rows, { inheritFrom })`, or after the last row with `appendRows(rows)`.
  Both take an array, an iterable or an async iterable, and read it only as the output drains, so a generator streams and nothing is held.
- Add a sheet with `editor.addWorksheet(name)`.
- Written files ask the application to recalculate when it opens them, so a total over cells you changed comes out right rather than showing the result the template was saved with.
- Writing a date into a cell that has no date format derives one and keeps the cell's font, fill and border.
- Convert between cell references and coordinates with `cellReference`, `columnIndexOf`, `columnLetters` and `rowNumberOf`.

Worth knowing before you rely on it.
Only `.xlsx` can be written, not `.ods`.
Rows are written over, never inserted, and nothing is pushed down, so a template with a totals block under its data region can have that block overwritten.
`save()` streams rather than returning bytes, so a failure part way through leaves an incomplete file, and calling it twice throws because your row sources have already been read.
A file past 4 GB, or with a single part past 4 GB, needs Zip64 and throws rather than being written wrong, the same way reading a Zip64 archive throws.


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

[Unreleased]: https://github.com/christophgockel/very-good-spreadsheet/compare/0.3.0...HEAD
[0.3.0]: https://github.com/christophgockel/very-good-spreadsheet/compare/0.2.0...0.3.0
[0.2.0]: https://github.com/christophgockel/very-good-spreadsheet/compare/0.1.0...0.2.0
[0.1.0]: https://github.com/christophgockel/very-good-spreadsheet/releases/tag/0.1.0
