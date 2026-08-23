# Very Good Spreadsheet

Read Excel `.xlsx` and OpenDocument `.ods` spreadsheets in Node and the browser, without holding the whole sheet in memory.

**[Try it live](https://spreadsheet.verygoodsoftware.net)**.
Drop in an `.xlsx` or `.ods` and watch it stream in your browser.

[npm](https://www.npmjs.com/package/@very-good-software/spreadsheet) | [GitHub](https://github.com/christophgockel/very-good-spreadsheet)

This library reads a spreadsheet as a stream.
It streams the rows and does not keep them, so the memory it uses stays low even for a very large sheet.
The same code runs in Node and in the browser, because it uses only web standard features and nothing specific to Node.
It reads files, and it writes them.
Point it at a template someone designed in Excel, write to the names they gave their data, and everything else is copied across untouched.


## Why this library

`xlsx` (SheetJS) and `exceljs` already read and write Excel files, and they've done it for years.
If you need a format we don't handle, use them.
This isn't a pitch to rip them out.

The thing is, they were written a long time ago.
Back then Node had no `ReadableStream` and no `DecompressionStream`, and ESM and TypeScript were nowhere near where they are now.
So they still include their own stream code, come with a big dependency tree and build to CommonJS.
They can't really drop any of that without breaking everyone already using them.

There is also a difference in how we write.
Both of them read a file into a model and write that model back out, so what they do not understand is not in the model and does not come back.
We copy every part we are not editing across as its own bytes, which is why a chart, a pivot table or a macro survives a fill: we never look at it.

That is also what makes the template model work.
You write to a region the template's author named, the region ends up as tall as your data, and the rest of the sheet moves.
Every formula that read those rows moves with them, on that sheet and on every other.
`exceljs` has `insertRow`, but it shifts values, styles and defined names without touching formulas, so a total over rows 6 to 20 still covers rows 6 to 20 afterwards.
We would rather refuse to write a file than write one that looks right and adds up wrong, so where we find something we cannot move, `save` throws and tells you what it is.

The upshot is a split worth having.
Whoever owns the template does the layout, the formatting and the formulas in Excel, where they can see what they are doing.
Your code supplies data.
There is no styling API here and there is not going to be one.

We have the luxury of being able to skip most of that.
The library is one codebase on web standards, so the same code reads a file in Node and in the browser.
It has one runtime dependency.
The types come straight from the source.
Memory stays flat as the file grows, because we read the sheet as a stream instead of loading the whole thing.

There are a few catches.
We read `.xlsx` and `.ods` but only write `.xlsx`, and nothing more exotic than those.
And we're new, so we haven't run into the years of odd files the older libraries have.
If that's fine for what you're building, give it a try.


## Status

This is early.
It reads `.xlsx`, the format Excel has used since 2007, and `.ods`, the OpenDocument spreadsheet format.
`Workbook.open` picks the format from the file, so the same code reads both.
It does not read the older `.xls` format or `.xlsb` yet.
The reading API is usable and covers the common cell types, and writing covers `.xlsx`.
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


## Writing a file

Writing takes a workbook, collects your edits, and produces a new file.
The input is never touched.

Every part of the file you do not edit is copied across byte for byte.
That means charts, pivot tables, drawings, macros and formatting survive, because we never look at them.
This is the point of the whole design, and it is why filling a client's template works.


### Filling a template

<!-- example: write-fill-template.ts -->

```ts
import { createWriteStream, openAsBlob } from "node:fs";
import { Writable } from "node:stream";
import { date, Workbook } from "@very-good-software/spreadsheet";

// Open a template, fill in the parts that change, and write a new file. The
// template's own charts, formatting and formulas are copied across untouched.
const workbook = await Workbook.open(await openAsBlob("template.xlsx"));
const editor = workbook.edit();

const invoice = editor.worksheet("Invoice");
invoice.set("C3", "Acme Ltd");
invoice.set("C4", date("2026-03-01"));

invoice.writeRows(8, [
  ["Consulting", 12, 950],
  ["Expenses", 1, 240],
]);

await editor.save().pipeTo(Writable.toWeb(createWriteStream("invoice.xlsx")));
```

<!-- /example -->

The cells you write keep whatever formatting they already had.
`writeRows` writes over the rows at that position without pushing anything down, which suits a template with a pre-formatted data region waiting to be filled.


### Writing where the template says, not where you guessed

`writeRows(8, ...)` is fine when the template lives beside your code and the same person changes both.
It stops being fine when the template belongs to someone else.
They insert a row for a subtitle, the data region moves to row 9, and your code keeps writing to row 8 with nothing to warn you.

So a region can be addressed by the name its author gave it.
In Excel that is Formulas, Name Manager, or selecting the region and typing a name into the Name Box.
Excel keeps a name pointing at the right cells when rows move, which is exactly what a cell reference cannot do.

<!-- example: write-named-region.ts -->

```ts
import { createWriteStream, openAsBlob } from "node:fs";
import { Writable } from "node:stream";
import { Workbook } from "@very-good-software/spreadsheet";

// The template's author selected the data rows in Excel and named them "Lines".
// They also put a total underneath, styled the region, and never told anyone which
// row anything is on, because none of that is this code's business.
const workbook = await Workbook.open(await openAsBlob("invoice-template.xlsx"));
const editor = workbook.edit();

const lines = await invoiceLines();

// However many lines there are, the region ends up that tall. The total below it
// moves with the rows and keeps covering all of them.
editor.worksheet("Invoice").writeRegion("Lines", lines);

await editor.save().pipeTo(Writable.toWeb(createWriteStream("invoice.xlsx")));

async function invoiceLines(): Promise<(string | number)[][]> {
  return [
    ["Consulting", 12, 950],
    ["Expenses", 1, 240],
    ["Support", 3, 180],
  ];
}
```

<!-- /example -->

The region ends up exactly as tall as the data you give it, and the sheet moves around it.
More rows than it covers pushes everything below down. Fewer pulls everything below up.
This is the model you already have from every other templating tool: content below the hole stays below the hole, however big the hole turns out to be.

It is also the same operation you would do by hand in Excel with Insert and Delete, so it behaves the way you would expect it to.
A total written `=SUM(C9:C11)` over a region of rows 9 to 11 becomes `=SUM(C9:C13)` when five rows arrive, and `=SUM(C9:C9)` when one does.
Formulas on other sheets that read those rows move too, and so do merged ranges, conditional formats, data validations, hyperlinks, filters, page breaks, frozen panes, table extents and the region's own name.

Give it no rows at all and you get one empty, correctly styled row.
The region never shrinks past that, because a total written over it would have its whole range deleted and die.

`worksheet.writeRegion` finds the names that worksheet owns, then the workbook-wide ones.
`editor.writeRegion` finds only the workbook-wide ones, and writes into whichever sheet the name points at.
That mirrors Excel, where a name can belong to one sheet or to the whole file, and a sheet's own name wins.

An unknown name fails at the call.
So does a name that cannot be written into, and the error says which kind it is, whether that is a formula, a print area, a whole column or a reference Excel broke when someone deleted a sheet.
Being told a name is a print area is more use than being told it is missing.

### Tables, for when you do not know how many rows there are

A named region has to fit.
An Excel Table does not, and that is the reason to use one.

Select your data in Excel and press Insert, Table.
You get a named object with a header row, filter buttons and, if you want it, a totals row.
Its name works with `writeRegion` exactly like a named region's, except that the rows go between the header and the totals row, and the table grows if you give it more rows than it currently holds.

Growing is the part that earns it.
A total written as `=SUM(Sales[Amount])` refers to the table by name rather than to a range of rows, so it keeps covering everything however many rows arrive.
Write that same total as `=SUM(C9:C20)` and it will not.

A table with a totals row is fine. That row moves down with everything else.

The table's own range and its filter's move with it, so Excel keeps treating the new rows as part of the table.

### Formatting is the template's job

There is no API here for fonts, fills, borders, merges or column widths, and there is not going to be one.
Those decisions belong to whoever owns the template, made in Excel or LibreOffice where you can see what you are doing.
We think you will get a better result in ten minutes there than through any API we could offer.

The case that tests this is "overdue rows should be red".
The answer is a conditional format, set in the template against a column you write a value into.
We copy that part across without reading it, so the rule survives and applies itself to whatever you write.

If that turns out to be too thin, the next step is letting you name a cell style the template already defines.
That would still leave the template deciding what the style looks like.
Describing formatting from code is the thing we are not doing.

### Starting from nothing

There is no separate mode for building a file.
`Workbook.create()` gives you an empty workbook, and everything after that line is identical.

<!-- example: write-from-scratch.ts -->

```ts
import { createWriteStream } from "node:fs";
import { Writable } from "node:stream";
import { formula, Workbook } from "@very-good-software/spreadsheet";

// Nothing to open, so start from an empty workbook. Everything after the first
// line is the same as filling a template.
const editor = (await Workbook.create()).edit();
const sheet = editor.worksheet(0);

sheet.appendRows([
  ["Region", "Units"],
  ["North", 120],
  ["South", 340],
]);
sheet.set("A5", "Total");
sheet.set("B5", formula("SUM(B2:B3)"));

await editor.save().pipeTo(Writable.toWeb(createWriteStream("report.xlsx")));
```

<!-- /example -->


### What you can put in a cell

A number, a string, a boolean, `null`, `formula(...)`, or `date(...)`.

```ts
sheet.set("A1", 42);
sheet.set("A2", "text");
sheet.set("A3", true);
sheet.set("A4", null); // blanks the cell and keeps its formatting
sheet.set("A5", formula("SUM(A1:A3)"));
sheet.set("A6", date("2026-03-01"));
sheet.set("A7", date(2026, 3, 1, 14, 30)); // month is 1 to 12, not 0 to 11
```

There is no error value.
An error is something a formula produces, so write `formula("NA()")` rather than the text `#N/A`.

A formula is written without its cached result, and the file asks the application to recalculate when it opens.
So a total over cells you changed comes out right, rather than showing the number the template was saved with.

**A `Date` is not accepted, and that is deliberate.**
A cell holds a calendar date with no time zone, while a `Date` is an instant, so turning one into the other means picking a zone and either choice is wrong for somebody.
`new Date(2026, 2, 1)` is local midnight, which west of UTC is the last day of February, and that is a silently wrong number in a report.
So you say which day you mean, and `date` refuses anything that is not one: `date(2026, 2, 30)` is an error rather than the 2nd of March.

If you already hold a `Date` and want the calendar values it has in UTC, pass `instant.toISOString()`.
For its local values, pass its parts.


### Big exports stay small in memory

`appendRows` and `writeRows` pull rows as the output drains, not when you hand them over, so a generator streams and nothing is held.

`writeRegion` streams too, and its rows are counted as they are written rather than beforehand.
A million rows into a region finishes under a 150MB heap, the same as `appendRows`.

There is one shape where it cannot.
Whatever sits above the region goes out before its rows have been counted, so if something up there reads rows below the region, a summary block at the top of the sheet for instance, the rows have to be counted first and are held while that happens.
Nothing else about the output differs, and you will not notice unless you are writing a great many rows into a template of that shape.

`npm run benchmark:write` prints every path side by side.
Add `--cap=150` to see what is actually held, rather than what the runtime has not yet given back to the operating system.

<!-- example: write-large-export.ts -->

```ts
import { createWriteStream } from "node:fs";
import { Writable } from "node:stream";
import { Workbook } from "@very-good-software/spreadsheet";

// The rows are pulled as the output drains, so this holds one row at a time
// however many there are.
async function* everyOrder(): AsyncIterable<readonly [number, string, number]> {
  for (let id = 1; id <= 5_000_000; id += 1) {
    yield [id, `Order ${id}`, id * 1.5];
  }
}

const editor = (await Workbook.create()).edit();
editor.worksheet(0).appendRows(everyOrder());

await editor.save().pipeTo(Writable.toWeb(createWriteStream("orders.xlsx")));
```

<!-- /example -->

Because nothing is read until you save, an error inside your own generator surfaces from `save` rather than from `appendRows`.
Everything we can check ourselves, a bad cell reference or an unknown sheet, throws at the call instead.

`save()` returns a stream rather than bytes, so the output is never held either.
The cost is that a failure part way through leaves an incomplete file with no undo.
A zip is read from a directory at its end, so a truncated write is a file no reader will open rather than a subtly wrong spreadsheet.
Saving twice throws, because your row sources have already been read and the second file would quietly be missing rows.


## What it does not write

- `.ods`.
  Only `.xlsx` can be written.
- Moving rows for `writeRows` or `appendRows`.
  Those write where you say and push nothing around.
  Rows move only for `writeRegion`, where the region says which rows are the data.
- Deleting or renaming a sheet.
  Both ripple into everything that refers to them by name.
- Growing a chart's range to cover rows appended below it, since appending moves nothing for the range to follow.
- Growing an Excel Table you filled by row number, since nothing in that call says the rows belong to it.
  Address the table by name and it grows.
- Anything on a sheet whose rows are moving that we cannot move with them.
  Rather than leave it stale, `save` throws and names it.
  That means a formula spanning a range of sheets, a whole row reference, a sheet carrying an extension list, which is where sparklines and the newer conditional formats live, rows that would take the whole of a pivot table's source range with them, and a pivot built from consolidation ranges rather than from one range.

Charts and images do move.
A shape anchored below the region comes down or up with the rows, and one anchored across the region stretches, the same as Excel does when you insert rows by hand.
The one case that stops a save is a shape standing only on rows that are going away, since there would be nothing left to hang it from.

A chart also moves the range each of its series reads, so a chart over a region you filled plots the rows you wrote rather than the cells they used to sit in.
That holds wherever the chart is drawn, on the sheet that moved, on another sheet, or on a tab of its own, because a series names the sheet it reads.
A chart built from a named range needs nothing moved, since the name moves itself.
A chart also carries its own copy of the values it read, which we do not rewrite, and Excel plots from the range rather than from that copy, so the figures are right on open with nothing to refresh.

Cell comments move too, both the cell they are attached to and the box they appear in.
A comment on a row that goes away goes with it, which is what Excel does when you delete that row by hand.

Pivot tables move as well.
The range a pivot reads follows the rows, and so does the pivot itself when it sits on the sheet that moved.
A pivot keeps its own copy of the source rows, which we do not rewrite, so the file asks Excel to rebuild it on open instead.
That means the figures are right once Excel has opened the file, and a reader that does not refresh caches shows the old ones.

- Keeping a digital signature valid.
  Any change to a file invalidates it.
- Files past 4 GB, or with a single part past 4 GB.
  Those need Zip64, which we do not write, and we throw rather than produce a file no reader will open.
  This mirrors the read side, which throws on a Zip64 archive.
  Handling files that large is its own piece of work.

Two things worth knowing about `writeRegion`.

Only one region per worksheet per save.
Writing one moves the rows the others were aimed at.

And a template with a summary block above the region reading rows below it is the one shape where the rows have to be counted before the file is written, rather than as it is written.
See "Big exports stay small in memory".


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
npm run verify        # lint, typecheck, and tests
npm run build         # build the package into dist
npm run benchmark     # reading and writing, against other libraries, see benchmark/
npm run benchmark:read   # reading only
npm run benchmark:write  # writing only, a million rows
npm run manual-check  # build a filled template to open in Excel, see MANUAL-CHECKS.md
```
