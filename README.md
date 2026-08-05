# Very Good Spreadsheet

Read Excel `.xlsx` and OpenDocument `.ods` spreadsheets in Node and the browser, without holding the whole sheet in memory.

**[Try it live](https://spreadsheet.verygoodsoftware.net)**. Drop in an `.xlsx` or `.ods` and watch it stream in your browser.

[npm](https://www.npmjs.com/package/@very-good-software/spreadsheet) | [GitHub](https://github.com/christophgockel/very-good-spreadsheet)

This library reads a spreadsheet as a stream.
It streams the rows and does not keep them, so the memory it uses stays low even for a very large sheet.
The same code runs in Node and in the browser, because it uses only web standard features and nothing specific to Node.
Reading is the only thing it does today.
Writing is planned for later.


## Why this library

`xlsx` (SheetJS) and `exceljs` already read Excel files, and they've done it for years.
If you need to write files, or read a format we don't, use them.
This isn't a pitch to rip them out.

The thing is, they were written a long time ago.
Back then Node had no `ReadableStream` and no `DecompressionStream`, and ESM and TypeScript were nowhere near where they are now.
So they still include their own stream code, come with a big dependency tree and build to CommonJS.
They can't really drop any of that without breaking everyone already using them.

We have the luxury of being able to skip most of that.
The library is one codebase on web standards, so the same code reads a file in Node and in the browser.
It has one runtime dependency.
The types come straight from the source.
Memory stays flat as the file grows, because we read the sheet as a stream instead of loading the whole thing.

There are a few catches.
We only read for now, no writing yet.
We handle `.xlsx` and `.ods` and nothing more exotic.
And we're new, so we haven't run into the years of odd files the older libraries have.
If that's fine for what you're building, give it a try.


## Status

This is early.
It reads `.xlsx`, the format Excel has used since 2007, and `.ods`, the OpenDocument spreadsheet format.
`Workbook.open` picks the format from the file, so the same code reads both.
It does not read the older `.xls` format or `.xlsb` yet.
The reading API is usable and covers the common cell types.
Expect the API to still change.

One `.ods` detail to know: a formula cell's `value` is the ODF formula text, which uses its own syntax, for example `of:=SUM([.B1:.B3])`.
The gaps in `.ods` reading are listed under "What it does not read".


## Install

```sh
npm install @very-good-software/spreadsheet
```

It needs Node 24 or newer and ships as an ES module.


## Reading a file

Open a workbook from its bytes, list its sheets, then stream a sheet's rows.

In Node:

<!-- example: read-node.ts -->

```ts
import { readFile } from "node:fs/promises";
import { Workbook } from "@very-good-software/spreadsheet";

const workbook = await Workbook.open(await readFile("data.xlsx"));

for (const name of workbook.worksheetNames) {
  console.log(name);
}

for await (const row of workbook.worksheet("Sheet1").rows()) {
  for (const cell of row.cells) {
    console.log(cell.ref, cell.type, cell.value);
  }
}
```

<!-- /example -->

In the browser, pass the chosen `File` straight in:

<!-- example: read-browser.ts -->

```ts
import { Workbook } from "@very-good-software/spreadsheet";

// A File from an <input type="file"> is a seekable Blob, so this reads it in
// ranges off disk instead of loading the whole file into memory first.
export async function readSpreadsheet(file: File): Promise<void> {
  const workbook = await Workbook.open(file);

  for await (const row of workbook.firstWorksheet().rows()) {
    console.log(row.number, row.cells);
  }
}
```

<!-- /example -->

`Workbook.open` accepts a `Uint8Array`, an `ArrayBuffer`, a `ReadableStream` of bytes, or a `Blob`.
Bytes and streams are read into memory and held.
A `Blob` is seekable, so the library reads it in ranges and never holds the whole file.
A browser `File` is already a `Blob`.
In Node you get a seekable `Blob` for a file on disk with `fs.openAsBlob`.

<!-- example: read-node-seekable.ts -->

```ts
import { openAsBlob } from "node:fs";
import { Workbook } from "@very-good-software/spreadsheet";

// openAsBlob gives a seekable Blob backed by the file on disk, so a large file
// is read in ranges and never held whole.
const workbook = await Workbook.open(await openAsBlob("data.xlsx"));

console.log(workbook.worksheetNames);
```

<!-- /example -->


## Cell values

Every cell has a type and a value that matches it.
The type is one of `number`, `string`, `boolean`, `date`, `error`, or `formula`.
You read the type first, and the value matches that type.

<!-- example: cell-values.ts -->

```ts
import type { Cell } from "@very-good-software/spreadsheet";

// Matching on a cell's type gives you a type-safe way to access its value.
export function show(cell: Cell): void {
  switch (cell.type) {
    case "number":
      console.log("number", cell.value);
      break;
    case "string":
      console.log("string", cell.value);
      break;
    case "boolean":
      console.log("boolean", cell.value);
      break;
    case "date":
      console.log("date, in UTC", cell.value);
      break;
    case "formula":
      // value is the formula text, cachedValue is the result Excel stored.
      console.log("formula", cell.value, cell.cachedValue?.value);
      break;
    case "error":
      console.log("error text", cell.value);
      break;
  }
}
```

<!-- /example -->

A date is stored in a spreadsheet as a number with a date format.
This library reads the styles and the workbook's date system, so a date cell comes back as a `Date`.

A formula cell has type `formula`.
Its `value` is the formula text as stored in the file, without a leading `=`.
Its `cachedValue` is the result Excel last computed, typed like any other cell, or `null` if the file holds no cached result.
The library does not recompute formulas.
Shared and array formulas are read as stored and not resolved, so a shared formula's dependent cells have an empty `value`.


## Rows and columns

`rows()` is an async iterable.
It yields a `Row` as each row is parsed.
A `Row` has its number, the cells that are present, and a lookup by column.

<!-- example: rows-and-columns.ts -->

```ts
import type { Worksheet } from "@very-good-software/spreadsheet";

export async function readGrid(sheet: Worksheet): Promise<void> {
  for await (const row of sheet.rows()) {
    const columnA = row.cell(0);
    console.log(row.number, columnA?.value, row.cells.length);
  }
}
```

<!-- /example -->

Empty cells are left out.
Each cell carries a `columnIndex`, a number starting at zero, and a `ref` like `"A1"`.
So you can place a row's cells into a fixed set of columns yourself, and decide how to fill an empty one.


## Hidden sheets

`workbook.worksheets` lists each sheet with its name and whether it is hidden, so you can skip hidden sheets.

<!-- example: hidden-sheets.ts -->

```ts
import type { Workbook } from "@very-good-software/spreadsheet";

export function visibleSheetNames(workbook: Workbook): string[] {
  return workbook.worksheets.filter((sheet) => !sheet.hidden).map((sheet) => sheet.name);
}
```

<!-- /example -->


## Reading all rows, or only a few

The API streams by default, and there is no `toArray` or `take`, because a plain loop already covers both.

To hold a whole workbook in memory, collect the rows yourself:

<!-- example: read-all-rows.ts -->

```ts
import type { Row, Workbook } from "@very-good-software/spreadsheet";

// There is no built-in "read everything" on purpose.
// Streaming keeps memory low by default, and holding the whole workbook
// is a plain loop when you want it.
export async function readAll(workbook: Workbook): Promise<Map<string, Row[]>> {
  const sheets = new Map<string, Row[]>();
  for (const name of workbook.worksheetNames) {
    const rows: Row[] = [];
    for await (const row of workbook.worksheet(name).rows()) {
      rows.push(row);
    }
    sheets.set(name, rows);
  }
  return sheets;
}

// Array.fromAsync(sheet.rows()) does the same for one sheet in a single line,
// where the runtime and your TypeScript lib provide it.
```

<!-- /example -->

To read only the start of a sheet, break out of the loop.
Because `rows()` is a stream, stopping the loop stops the reading, so a preview of a huge file only parses the rows you take:

<!-- example: first-ten-rows.ts -->

```ts
import type { Row, Worksheet } from "@very-good-software/spreadsheet";

// There is no take or limit on purpose. rows() is a stream, so breaking the loop
// stops the reading, and only the first ten rows of a huge file are ever parsed.
export async function firstTenRows(sheet: Worksheet): Promise<Row[]> {
  const preview: Row[] = [];
  for await (const row of sheet.rows()) {
    preview.push(row);
    if (preview.length === 10) {
      break;
    }
  }
  return preview;
}
```

<!-- /example -->


## Streaming and memory

The reason to use this library is memory.
Most libraries read a whole `.xlsx` into memory, expand every part, and build a full model of it.
That is fine for a small file.
For a large one it can use a lot of memory or run out of it.

When you pass bytes, the library holds the compressed file and a small working set, and nothing more.
It has to keep the compressed file, because a zip is addressed from a directory at its end, so a named part cannot be pulled from a forward only stream.
It decompresses each part as you pull it, through the platform's `DecompressionStream`, and parses the XML as a stream.
So when you iterate rows and let each one go, the uncompressed sheet is never held at once, and the rows are not kept.
Memory stays flat as the sheet gets longer.
It tracks the compressed file size and a fixed working set, not the row count or the uncompressed size.

When you pass a `Blob`, it does not hold the compressed file.
It reads the central directory and then each part in ranges off the source, so the file itself is never resident, only a window and the working set.
Reading the same file as bytes holds all of it, so a `Blob` lowers the memory you truly need by roughly the file size.
This saving grows with the file.
It is not visible on a file of a few tens of MB, and on a 130 MB file it lowered the held memory by about that file's size in our test.

Our benchmark file is 28 MB on disk, and its one sheet is 170 MB once decompressed, with 4.44 million cells.
Reading every cell peaks at about 145 MB, and that peak does not grow as the sheet gets longer.
The libraries that load everything use a few gigabytes on the same file.
Under a 512 MB memory limit, this library finishes and they run out of memory.

On xlsx it is fast too.
On that same file it streams faster than the other readers in our benchmark while using the least memory.
`.ods` is slower, because the whole document is one stream we re-read for each sheet.
But the memory stays just as bounded, and on a large `.ods` we finish where the load-everything readers run out.


## What it does not read

Some things are out of scope, and the library throws rather than guess when it reaches one of these:

- Zip64 archives, which very large files use.
- Encrypted or password protected files.
- A cell type it does not recognise.

In `.ods` files, some things read but not fully:

- Error cells read as their numeric fallback, because ODF keeps errors in an extension we do not read.
- Shared and nested tables are not resolved, so a shared formula's dependent cells read with an empty formula.
- A time cell reads as a string of its raw ISO duration, since the value types have no time.


## Development

```sh
npm run verify     # lint, typecheck, and tests
npm run build      # build the package into dist
npm run benchmark  # compare against other libraries, see benchmark/
```
