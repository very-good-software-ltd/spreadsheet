# Context

This file records the decisions behind this library and the questions still open.
It is a living document, not a set of immutable records.
When a decision changes, edit it here and note what changed.
If the project grows and a decision hardens, we can graduate it into a proper ADR later.


## What we are building

A modern library for reading Excel files.
The existing options like `xlsx` and `exceljs` feel dated or are not maintained often enough.
This is an honest attempt to see how far a fresh, modern take can get.


## Decisions

### 1. Format scope: `.xlsx` and `.ods`

We read `.xlsx` (OOXML) and `.ods` (OpenDocument), both a ZIP of XML parts, so they share the zip and xml layers and the `WorkbookData` seam.
`Workbook.open` sniffs the archive to pick the reader, `xl/workbook.xml` for xlsx and the spreadsheet `mimetype` for ods.
Legacy `.xls` (BIFF8, OLE2 compound binary) and `.xlsb` are still out of scope. They are separate binary containers with their own parsers.

Changed: `.ods` was originally out of scope. It was added once the format-neutral seam made a second ZIP based format cheap. The `.ods` reader has known gaps, see the open questions.


### 2. Reading first, then writing

Version 1 reads files.
The whole read design was shaped around reading alone, and writing was designed on top of it afterwards rather than alongside it.

Changed: writing was originally deferred with no design at all. Decisions 11 and 12 now cover it. The read path is unchanged by it, because the writer works below the `WorkbookData` seam on the archive and the raw XML, not on the interpreted rows.


### 3. Runs in Node and the browser

The core is runtime agnostic.
It is built on web standard primitives like `Uint8Array`, `ReadableStream`, and `DecompressionStream`.
Anything Node specific like `fs` lives in a thin adapter at the edge, not in the core.
We ship and test Node first, but browser support should be a packaging exercise, not a rewrite.


### 4. Third party libraries behind interfaces we control

Each backing library sits behind an interface we own, so we can swap the implementation or hand roll one without touching callers.
`saxes` backs the XML interface.
The ZIP interface started on `fflate` and now uses a hand rolled reader (see decision 9), which is exactly the swap the interface was built to allow.
`fflate` stays as a development dependency, used in tests as an independent writer to check our reader.


### 5. The XML interface is streaming shaped

The XML interface exposes a stream of pull events (`open`, `text`, `close`) over a `ReadableStream`, not a whole parsed tree.
A tree shaped interface would quietly lock us into loading everything into memory and would break the moment we swap in a streaming parser.
The event shape is satisfiable by a tree based library today and by a hand rolled streaming parser tomorrow.
The ZIP interface follows the same idea: list entries, open an entry as a stream.


### 6. The public API is streaming first

The primary API is async iterables for sheets and rows.
This is to better support large files and matches the engine underneath.
We add no eager helper like `.toArray()`. The platform's `Array.fromAsync` or a plain loop collects everything when a caller wants it, and breaking the loop reads only the first rows.


### 7. TypeScript 7, strict, Vite and Vitest

Authored in TypeScript 7 with strict settings.
See `tsconfig.json` for the exact flags.
Built and tested with Vite and Vitest.
Vitest can run the same suite in both Node and a real browser, which directly de-risks the part most likely to bite us, whether our stream and decompression code behaves the same in both runtimes.

We deliberately do not use `isolatedDeclarations`.
It only speeds up type file generation, which we do not feel at this size, and it taxes every exported symbol with an explicit type annotation.

One `tsconfig.json` covers the whole repo.
A split earns its place only when the repo holds files that run in environments with contradictory global types, for example browser code that must not see Node globals.
Our source uses only web standard globals, so there is one world and one config.
Revisit when we add a Node only adapter alongside a browser only path.


### 8. Shipping the library

We publish to npm.
`npm publish` uploads `dist/`, which holds the bundled ES module and the generated type files.
Vite builds the JavaScript and TypeScript 7's own `tsc` emits the type files, so the whole toolchain runs on one TypeScript version with no second compiler.
ES module only, no CommonJS build.

The type emit needs its own `tsconfig.build.json`. The root `tsconfig.json` sets `noEmit` and covers `src`, `test`, and `examples` for checking, so the build config extends it, narrows the input to `src`, and turns emit on. This is the one config split decision 7 allows for a reason other than global types, an emit that must see only the shipped source.

We do not use a dts plugin. `unplugin-dts` needs a 6.x compiler for the JavaScript Compiler API that TypeScript 7's native port dropped, and `rolldown-plugin-dts`'s tsgo backend, which would use TypeScript 7, is experimental and its bundler fails on our declarations with a false missing-export error. Plain `tsc` emit avoids both and produces the same per file declarations.


### 9. A hand rolled ZIP reader for real streaming

The ZIP reader parses the central directory itself and decompresses each entry on demand through the platform's `DecompressionStream`.
This keeps memory bounded for a large sheet, since the decompressed bytes are pulled through the parser and never fully held.
`fflate` could not give us backpressured decompression, so it moved to a development dependency.
We accept the maintenance of hand rolled binary parsing because the part of the ZIP format that xlsx uses is frozen, the surface is small and read only, and correctness is pinned by comparing our reader against `fflate` on many archives.
Zip64, encryption, and unknown compression methods are not supported and throw rather than misread.


### 10. Honest about size and memory claims

Size and runtime memory will grow as we add features like writing, more formats, and a formula engine, so we do not lean on an absolute "tiny" in evergreen copy.
The durable claim is the read path's memory, flat as the file grows, because it streams rows instead of holding the sheet.
Opt-in heavy features like formula evaluation will not be flat, so we do not claim the whole library stays lean.
Specific size and peak-memory numbers live in the benchmark, where they are dated and reproducible, not in taglines.

Changed: this used to also frame bundle size as paying only for what you use, on the condition that reading never pulls in the writer.
Decision 12 gives that up on purpose. `edit()` is an instance method on `Workbook`, which is a hard reference no bundler can drop, so every reader now ships the writer.
We chose one coherent entry surface over a tree-shakeable one. The claim is dropped rather than quietly falsified.


### 11. Writing copies the file and rewrites only what it must

Writing takes a base workbook, collects edits, and produces a new file in one streaming pass at save.
Nothing is written per cell, and nothing is modified in place. The input is never touched.

Every archive entry we do not edit is copied across as stored bytes, without recompressing and without parsing.
Charts, pivot tables, drawings, themes, macros and custom parts survive because we never look at them.
This is the whole point. `exceljs` reads a file into a model and re-serialises the whole thing from an allowlist of parts it understands (`lib/xlsx/xlsx.js`), so read-then-write drops pivot caches, chart parts and custom XML. A copy-through writer cannot lose a part it does not know about.

For the same reason the sheet rewrite is a transform over the XML event stream, not a serialisation of our `Cell` model.
Our read model is deliberately lossy. It drops the style index, merged ranges, column widths, data validations and conditional formats.
Rebuilding a sheet from it would destroy exactly the formatting a template exists to carry.
So we pass every event through untouched and intervene only at the cells being written.

Creating a file from scratch is the same path over an empty base workbook, so there is no separate write mode and no second code path.

Scope is `.xlsx` only. `.ods` writing is a second full implementation that shares none of the fidelity work.

Rows can be written at a given row number, overwriting what is there, or appended after the last row.
Inserting rows and shifting content down is out.
Shifting content means rewriting every formula, merged range, conditional format and table range that pointed below the insertion.
`exceljs` offers `insertRow` and `spliceRows` and does not do this: it shifts values and styles and adjusts defined names, but never touches formulas, so a total that summed rows 6 to 20 still sums 6 to 20 afterwards.
Silently wrong numbers in a report is a worse failure than a missing feature, so we do not offer the feature.

### 12. The write API is one surface with the reader

```
Workbook.open(source)  → Promise<Workbook>
Workbook.create()      → Workbook            // blank, one sheet named Sheet1
workbook.edit()        → Editor
editor.worksheet(nameOrIndex)
  .set(ref, value)
  .writeRows(startRow, rows, { inheritFrom })
  .appendRows(rows)
editor.addWorksheet(name)
editor.save()          → ReadableStream<Uint8Array>
```

Two statics both hand back a `Workbook`, and one instance method flips to writing, so from-scratch reads as what it is, an empty template.

`save` returns a stream rather than bytes, and rows are pulled from their source as it drains.
Bytes would mean holding the whole output, which breaks on the way out the memory promise the read path keeps on the way in.
The cost is that a failed save leaves a partial file with no undo. We accept it because a zip's directory sits at its end, so a truncated write is a file no reader will open rather than a subtly wrong spreadsheet.
`save` throws if called twice, because an already-drained row source would silently produce a valid file with rows missing.

Row sources are `Iterable` or `AsyncIterable`, so an array works with no ceremony and a generator streams.
Nothing is consumed until save, which means an error in a caller's generator surfaces from `save` and not from the call that appeared to do the work.
Everything else we can check, a malformed reference or an unknown sheet, is validated eagerly at the call, so the caller's own source is the only thing left that can fail late.

A cell value is `number | string | boolean | Date | null | Formula`.
Errors are not writable. An error is a result, not an input, and the way to put one in a cell is a formula that produces it, such as `NA()`.
`null` blanks a cell and keeps its formatting, which is the template intent. Removing a cell outright is not offered.
Writing a plain value over a formula cell drops the formula, which is the one place the writer discards template content on purpose.
`formula()` is a constructor rather than a bare string so that ods writing could later translate inside it without touching a call site. Its text is the A1 form Excel uses, a leading `=` is stripped, and it is emitted unparsed and unvalidated.

A cell is addressed by its reference string, matching `Cell.ref` on the read side.
Numbers would inherit an inconsistency we already have, since `Row.number` is one-based from the file while `Cell.columnIndex` is zero-based.
A coordinate-to-reference helper can be added if callers want one, which also answers the column letter question below.

Last write wins when a cell is covered twice.
Rejecting a conflict is not implementable, since spotting the overlap would mean draining the lazy row source at call time.

Deleting and renaming sheets is out, for the same reason inserting rows is: both ripple into everything that references them by name or id.


## Open questions

- **Cell value typing (resolved).**
  A cell is a typed union: number, string, boolean, date, error, and formula. The caller reads `type` and gets a matching `value`, rather than raw strings to interpret. A formula cell carries its text as `value` and its cached result as `cachedValue`, and `ResolvedValue` is the union minus the formula variant, so a cached result cannot itself be a formula.
- **Scope of content for v1 (resolved).**
  Formulas, styles, number formats, and date interpretation are all in, not raw values only. A date comes back as a `Date`, from the serial number and its style in xlsx or the ISO value in ods. A formula reads as its cached value.
- **Shared strings and large sheets (resolved).**
  The sheet XML streams row by row and is never held. The shared strings table is read into memory and held, because cells reference it by index. So peak tracks the shared strings table plus a working window, not the sheet size, and a huge shared strings table is the remaining floor.
- **Column identifier: number vs letter.**
  A cell's column is a zero-based `columnIndex` for now.
  We considered the spreadsheet letter (`"A"`, `"AB"`) since it feels closer to Excel, but consumers do positional and arithmetic work with columns, and the letter is already in the cell `ref`.
  Punted, revisit if consumers want letter-native access.
  We could also expose the letter alongside the index later.
- **Streaming throughput (resolved).**
  Profiling the 170MB read showed the cost was yielding one XML event at a time through async generators, not saxes, which parsed the whole sheet in about 2.2s.
  Two changes fixed it.
  The XML reader now yields the events from one chunk as a batch, so the async boundary is crossed once per chunk.
  And the reader holds saxes's attributes object instead of copying it on every open tag.
  Together these took the streaming time from about 9.9s to 3.4s, which is the fastest and the leanest on memory of the libraries we compared.
- **Browser packaging.**
  The library ships one ES module built on web standards (`Uint8Array`, `ReadableStream`, `DecompressionStream`, `Blob`), so it runs in the browser as is, and the demo site reads a `File` there. Still open: running the library's own test suite in a real browser, not just Node, to confirm the stream and decompression code behaves the same.
- **String encoding when writing (resolved).**
  A written string cell is embedded inline (`inlineStr`), never added to the shared strings table.
  Shared strings dedupe and match what Excel itself emits, so files are smaller, but the writer would have to hold the whole table to dedupe against it, which fights the single streaming pass, and filling a template would mean rewriting a part we would otherwise copy through untouched.
  Inline strings need no global table, so a sheet writes in one pass and the shared strings part stays byte-identical to the source.
  The cost is a larger file when the same string repeats. Reading handles both, so nothing downstream cares.
  Offering shared strings as an option later is possible, but it would need the table held in memory and is not worth it until someone asks.
- **Seekable sources for lower peak memory (resolved).**
  Passing bytes reads the whole input into memory and holds it for the life of the workbook, because a zip is addressed from a directory at its end, so a named part cannot be pulled from a forward only stream.
  `Workbook.open` now also accepts a `Blob`, a `File` in the browser or `fs.openAsBlob` in Node, read through a `ByteRange` in ranges, so the compressed file is never held.
  This lives at the IO and zip layer, below the `WorkbookData` seam, so it did not change the format readers, and every zip based format shares it through the `ZipArchive` interface.
  The saving is roughly the file size in the memory actually needed, and it grows with the file, so it is invisible on a small file and clear on a large one.
  It is not a flat cap on peak, since the reported peak also carries memory the runtime frees but does not return to the OS.
  `.xls` is a different container and would need its own range reader.
- **`.ods` reader gaps.**
  Error cells live in a `calcext` extension namespace we do not read, so an errored formula reads as its numeric fallback.
  Shared and nested tables are not resolved, so a shared formula's dependent cell reads with an empty formula text.
  A time cell reads as a string of its raw ISO duration, since the value types have no time.
  ODF keeps the whole spreadsheet in one `content.xml`, so listing sheets and reading a sheet each scan that part, unlike xlsx's per sheet parts. The list scan runs once at open.
- **`.ods` reading speed (profiled, deferred).**
  ODF keeps every sheet in one `content.xml` deflate stream, which cannot be seeked, so each access re-decompresses and re-parses from the start. On a 14 MB file that is 359 MB uncompressed: listing sheet names is about 3.2s (the XML reader tokenises cells it then discards, decompression alone is only about 0.5s), and reading a sheet re-parses from the start up to its table, up to 3.4s for the last sheet. Reading every sheet is about 25s, roughly one full parse per sheet.
  Profiled and left for now. A targeted byte scan for the `<table:table>` tags would cut the name list to about 0.5s, and skipping past earlier sheets would cut a single sheet read, but both trade the XML reader for hand rolled scanning, and reading every sheet at once is not a real use case. The memory story is unaffected and stays bounded. Revisit when a real workload hits it.
- **Website presents both formats (resolved).**
  The site covers `.ods` everywhere it matters now: the hero tagline, the page meta description, the demo copy, and a dedicated `.ods` block in the Benchmark section framed on the memory tradeoff (`ODS_BENCHMARK` in `website/app/site.ts`). The xlsx benchmark is labelled "Reading xlsx" so it reads as one of two scenarios, and the "Fast" feature blurb is scoped to xlsx so it does not contradict the slower `.ods` numbers.
- **Formula text is format-specific, and the read side leaks it.**
  The same formula is `SUM(A1:B1)` in xlsx and `of:=SUM([.A1:.B1])` in ods. That is a namespace prefix, a leading `=` and a different reference syntax, not a punctuation difference.
  Reading hands out each format's text as a plain string, so nothing stops a caller reading an ods formula and writing it into an xlsx as literal nonsense.
  The writer does not cause this, it makes it reachable. A format-neutral formula needs a parser that rewrites references and maps function names, which is a smaller job than the formula engine but still a real one.
  `formula()` exists as the boundary where that translation would go.
- **How faithful is a rewritten part.**
  Untouched entries are copied as bytes and are exactly identical. A part we rewrite is re-emitted from the XML event stream, which does not carry attribute order, self-closing tag spelling, comments or processing instructions.
  So the rewritten sheet is semantically equivalent but not byte-identical, and the byte-identity test can only cover the parts we did not touch. Whether that gap ever matters is unknown.
- **Recalculation on open needs confirming against real Excel.**
  Changing a value makes every cached formula result downstream of it stale.
  The plan is to drop `xl/calcChain.xml`, which is a cache Excel rebuilds, and set `fullCalcOnLoad` so Excel recomputes on open.
  Dropping a part means also removing its content-type override and its relationship, or confirming Excel tolerates the dangling ones. Neither is verified yet.
  The `dimension` element is a related unknown. It sits before the row data, so a single pass cannot know the final extent when it has to be emitted. It is optional in the schema, so dropping it should be safe, and that should be checked rather than assumed.
- **What a written file does not update.**
  An Excel Table in the template does not grow to cover appended rows, and a chart pointing at a fixed range does not extend.
  A digitally signed workbook has its signature invalidated by any modification, which is inherent and not fixable.
  A template whose data region sits above a totals block is not served: writing past the region overwrites the totals, and we cannot detect it because we do not know those rows are a totals block.
- **Verifying the writer needs real files.**
  The headline test is that every entry we did not edit is byte-identical to the source, since that is the promise stated mechanically.
  `exceljs` and `xlsx` read the output back as independent checks, the same way they already back the read tests.
  Neither catches Excel offering to repair a file, which is the failure that matters most, so that stays a manual check at release.
  All of this needs real template files with genuine cruft in them, charts and merged cells and conditional formats. Files we generate ourselves would pass every test and prove nothing.
- **Writing `.ods`.**
  Out of scope. It shares the zip and XML writing layers and nothing else, since the fidelity work is per format.
  It would also force the formula translation question above, because `formula()` cannot pass its text through untranslated.
