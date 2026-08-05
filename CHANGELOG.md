# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [Unreleased]

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
