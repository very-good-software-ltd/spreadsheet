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
Legacy `.xls` (BIFF8, OLE2 compound binary) and `.xlsb` are still out of scope.
They are separate binary containers with their own parsers.

Changed: `.ods` was originally out of scope.
It was added once the format-neutral seam made a second ZIP based format cheap.
The `.ods` reader has known gaps, see the open questions.


### 2. Reading first, then writing

Version 1 reads files.
The whole read design was shaped around reading alone, and writing was designed on top of it afterwards rather than alongside it.

Changed: writing was originally deferred with no design at all.
Decisions 11 and 12 now cover it.
The read path is unchanged by it, because the writer works below the `WorkbookData` seam on the archive and the raw XML, not on the interpreted rows.


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
We add no eager helper like `.toArray()`.
The platform's `Array.fromAsync` or a plain loop collects everything when a caller wants it, and breaking the loop reads only the first rows.


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

The type emit needs its own `tsconfig.build.json`.
The root `tsconfig.json` sets `noEmit` and covers `src`, `test`, and `examples` for checking, so the build config extends it, narrows the input to `src`, and turns emit on.
This is the one config split decision 7 allows for a reason other than global types, an emit that must see only the shipped source.

We do not use a dts plugin.
`unplugin-dts` needs a 6.x compiler for the JavaScript Compiler API that TypeScript 7's native port dropped, and `rolldown-plugin-dts`'s tsgo backend, which would use TypeScript 7, is experimental and its bundler fails on our declarations with a false missing-export error.
Plain `tsc` emit avoids both and produces the same per file declarations.


### 9. A hand rolled ZIP reader for real streaming

The ZIP reader parses the central directory itself and decompresses each entry on demand through the platform's `DecompressionStream`.
This keeps memory bounded for a large sheet, since the decompressed bytes are pulled through the parser and never fully held.
`fflate` could not give us backpressured decompression, so it moved to a development dependency.
We accept the maintenance of hand rolled binary parsing because the part of the ZIP format that xlsx uses is frozen, the surface is small and read only, and correctness is pinned by comparing our reader against `fflate` on many archives.
Zip64, encryption, and unknown compression methods are not supported and throw rather than misread.
The writer added by decision 11 holds the same line from the other side.
A size, an offset or an entry count past what a 32-bit or 16-bit header field holds throws rather than writing the wrapped value, which would produce an archive no reader opens.
So 4 GB is the ceiling in both directions, and lifting it is one piece of work on both halves rather than two.


### 10. Honest about size and memory claims

Size and runtime memory will grow as we add features like writing, more formats, and a formula engine, so we do not lean on an absolute "tiny" in evergreen copy.
The durable claim is the read path's memory, flat as the file grows, because it streams rows instead of holding the sheet.
Opt-in heavy features like formula evaluation will not be flat, so we do not claim the whole library stays lean.
Specific size and peak-memory numbers live in the benchmark, where they are dated and reproducible, not in taglines.

Changed: this used to also frame bundle size as paying only for what you use, on the condition that reading never pulls in the writer.
Decision 12 gives that up on purpose.
`edit()` is an instance method on `Workbook`, which is a hard reference no bundler can drop, so every reader now ships the writer.
We chose one coherent entry surface over a tree-shakeable one.
The claim is dropped rather than quietly falsified.


### 11. Writing copies the file and rewrites only what it must

Writing takes a base workbook, collects edits, and produces a new file in one streaming pass at save.
Nothing is written per cell, and nothing is modified in place.
The input is never touched.

Every archive entry we do not edit is copied across as stored bytes, without recompressing and without parsing.
Charts, pivot tables, drawings, themes, macros and custom parts survive because we never look at them.
This is the whole point.
`exceljs` reads a file into a model and re-serialises the whole thing from an allowlist of parts it understands (`lib/xlsx/xlsx.js`), so read-then-write drops pivot caches, chart parts and custom XML.
A copy-through writer cannot lose a part it does not know about.

For the same reason the sheet rewrite is a transform over the XML event stream, not a serialisation of our `Cell` model.
Our read model is deliberately lossy.
It drops the style index, merged ranges, column widths, data validations and conditional formats.
Rebuilding a sheet from it would destroy exactly the formatting a template exists to carry.
So we pass every event through untouched and intervene only at the cells being written.

Creating a file from scratch is the same path over an empty base workbook, so there is no separate write mode and no second code path.

Scope is `.xlsx` only.
`.ods` writing is a second full implementation that shares none of the fidelity work.

Rows can be written at a given row number, overwriting what is there, or appended after the last row.
Neither moves anything, and neither ever will.
`writeRows` and `appendRows` are the streaming path, where the caller knows the shape of the sheet and nothing is below the rows being written.

Changed: this used to refuse row insertion outright, for the reasons decision 15 now records and answers.
Moving content lives there, on a named region, and not on these two.


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
The cost is that a failed save leaves a partial file with no undo.
We accept it because a zip's directory sits at its end, so a truncated write is a file no reader will open rather than a subtly wrong spreadsheet.
`save` throws if called twice, because an already-drained row source would silently produce a valid file with rows missing.

Row sources are `Iterable` or `AsyncIterable`, so an array works with no ceremony and a generator streams.
Nothing is consumed until save, which means an error in a caller's generator surfaces from `save` and not from the call that appeared to do the work.
Everything else we can check, a malformed reference or an unknown sheet, is validated eagerly at the call, so the caller's own source is the only thing left that can fail late.

A cell value is `number | string | boolean | null | Formula | SpreadsheetDate`.
Errors are not writable.
An error is a result, not an input, and the way to put one in a cell is a formula that produces it, such as `NA()`.
`null` blanks a cell and keeps its formatting, which is the template intent.
Removing a cell outright is not offered.
Writing a plain value over a formula cell drops the formula, which is the one place the writer discards template content on purpose.
`formula()` is a constructor rather than a bare string so that ods writing could later translate inside it without touching a call site.
Its text is the A1 form Excel uses, a leading `=` is stripped, and it is emitted unparsed and unvalidated.

`date()` is a constructor for the same reason, and a `Date` is refused outright.
A cell holds a calendar date with no time zone while a `Date` is an instant, so converting one means picking a zone, and either choice is silently wrong for some caller.
`new Date(2026, 2, 1)` is local midnight, which west of UTC is the last day of February, and date libraries hand out local-midnight `Date`s by default.
We first shipped this accepting a `Date` read as UTC and documenting the trap.
That failed our own bar: a confidently wrong quirk is worse than an open one, and a wrong day in a report is the worst failure class here.
Refusing the type moves the mistake to compile time.
`date()` also rejects a day that does not exist instead of letting `Date.UTC` roll it into the next month.

A cell is addressed by its reference string, matching `Cell.ref` on the read side.
Numbers would inherit an inconsistency we already have, since `Row.number` is one-based from the file while `Cell.columnIndex` is zero-based.
A coordinate-to-reference helper can be added if callers want one, which also answers the column letter question below.

Last write wins when a cell is covered twice.
Rejecting a conflict is not implementable, since spotting the overlap would mean draining the lazy row source at call time.

Deleting and renaming sheets is out.
A sheet's name is written into every formula that reads it, into defined names, into pivot sources and into chart series, and its id into the parts that point at it, so either one ripples through the whole package rather than through one sheet.


### 13. Formatting is described in the template, never in our API

The writer will never grow a way to describe formatting.
No fonts, fills, borders, merges, column widths or number formats.
Those decisions belong to whoever owns the template, made visually in Excel or LibreOffice, where the feedback is immediate and the tool is better than any API we could offer.

The pressure on this never arrives as a request for a styling API.
It arrives as "overdue rows should be red", which is reasonable and which every other library answers with a formatting object.
Our answer is conditional formatting in the template, driven by a value we write.
The author sets the rule in Excel against a column we fill, so the rule survives untouched because we copy the part through without reading it.

The escape hatch we will build if that proves too thin is pointing at a named cell style the template already defines, an entry in `styles.xml` `cellStyles` that the author created in Excel.
That adds a reference, not a vocabulary, so the template still owns what the style means and we only own which one applies.
It is deliberately not in the first version, because we want to find out whether conditional formatting actually suffices before adding surface we cannot remove.

The line, stated positively: the library lets a caller point at formatting the template already has, and never lets them describe formatting of their own.

Changed: this used to be an absence rather than a decision.
The writer shipped with no styling surface because none was needed yet, which reads the same from outside as a gap waiting to be filled.
It is now a position.


### 14. Writing is anchored to what the template names, not only to coordinates

There are two kinds of template and we serve both.
One is a developer artifact, checked in beside the code, changed in the same commit as the call site.
Coordinates are fine there, and `set("C3", value)` is no worse than any other constant.
The other is a business artifact, owned by whoever needs the report, edited in Excel without telling anyone.
There a coordinate is a silent wrong-cell bug the first time someone inserts a row for a subtitle.

So a caller can also address a named region, and the name is the contract between the template's author and the code.
Excel maintains a name across insertion and deletion, which is exactly the case a coordinate cannot survive.
The vocabulary is ours and neutral, a named region of cells, not the format's word for it, so ODF's named expressions can back the same API later.
This is the same reason `formula()` is a constructor.

The order of anchors is deliberate.
Named regions come first, because the author placed them on purpose, the format keeps them correct, and they cost one more element in `xl/workbook.xml`, a part we already stream at open.
Excel Tables come second.
A table's writable region is the rows between its header and its totals row, since neither is a place for a caller's data, while its extent covers both.

Its extent and its autofilter range move with the rows, so a total written as a structured reference like `=SUM(Sales[Amount])`, or as a whole column `=SUM(B:B)`, keeps covering everything however many rows arrive.
Neither names a row number, which is the whole reason to anchor on a table.

Changed: a table with a totals row used to be refused, and a table used to grow but never shrink.
Both existed only because nothing could move.
Decision 15 moves rows, so the totals row goes down with everything else and a table pulls in as readily as it stretches.
Matching a header row by its text does not happen.
It is a guess wearing an anchor's clothes, and a guess that silently picks the wrong row is the failure class we turned down `insertRow` over.

A named region is a contract about an area, not a starting coordinate.
A row wider than the region is refused rather than spilling into the column beside it.
A row that stops short has the rest of its columns cleared, because what is sitting there is the last run's data formatted exactly like this run's, while a gap the caller writes as `undefined` still means leave that cell alone, since the length of a row is not a decision about the columns past its end.

Changed: the region's height used to be a contract too, so more rows than it held was refused and fewer cleared the remainder in place.
Decision 15 replaces that.
The region is now as tall as the data and the sheet moves around it.
The width is still a contract, because nothing moves sideways.
This is what closes the totals-block gap in the open questions below.
We could not detect a totals block because we did not know where the data region ended.
A named region is the author telling us.

The rigidity is the point, and it is also the cost.
A region that must fit exactly will not suit a caller whose data varies in length.
That caller wants `appendRows`, which promises nothing about what is below, or an Excel Table once those land.

`<definedNames>` holds far more than named boxes, so only part of it is addressable: a single area, absolute, in this workbook, pointing at cells, and not one of Excel's own `_xlnm.` entries.
A name that is a constant, a formula, a multi-area range, a relative reference, an external workbook reference or a `#REF!` left behind by a deleted sheet throws at the call and says which of those it is.
A whole column name like `Sheet1!$B:$B` throws too, since clearing the remainder of a region a million rows tall is not a thing to attempt.
Print areas are refused along with the other built-ins, which we can revisit if anyone asks.
A name is either global or scoped to one sheet, and a sheet-scoped name shadows a global one, so the sheet's editor resolves its own names first and the workbook's editor sees only global ones.
The shadowing is visible in the API rather than a rule to remember.


### 15. A named region holds exactly the data it is given, and the sheet moves around it

Changed: decision 11 refused row insertion outright, and decision 14 made a named region a fixed box.
Both are replaced here.
What follows is why.

A template is a document with a hole in it.
Everything above the hole stays, everything below it stays below, and the hole is as big as the data.
That is what every other templating tool does and it is the model a template author already has, so a library that cannot do it has to teach a workaround instead: put the total above the data, do not park anything under the region, use a table if you want it to grow.
Each of those is a thing we would rather not have to say.

So writing into a named region makes the region exactly as tall as the rows it is given.
More rows than it holds inserts rows and pushes everything below down.
Fewer deletes rows and pulls everything below up.
This is the same operation a person performs in Excel with Insert and Delete, so the semantics are not ours to invent.
They are observable, and a disagreement with Excel is a bug rather than a design question.

Where the rows go in is not arbitrary.
Excel stretches a range only when rows appear strictly inside it, so making room happens at the region's last row rather than after it.
A total written `=SUM(C9:C11)` over a region of rows 9 to 11 then becomes `=SUM(C9:C13)` when two rows arrive, which is what the author meant.
Inserting after row 11 would leave it summing three rows of five.
Deleting works from the far end for the same reason, so the first rows and their formatting are the ones kept.

Decision 11's objection stands as a description of the work rather than as a reason not to do it.
Shifting content means rewriting everything that pointed below the insertion, and `exceljs` offers insertion without doing that, which is why a total that summed rows 6 to 20 still sums 6 to 20 afterwards.
Silently wrong numbers are still worse than a missing feature.
The conclusion changes because there is a third option we did not consider: do the work where we can and refuse the file where we cannot.

So a save either produces a correct file or throws naming what stopped it.
Never a file that is quietly wrong.

In scope to move: rows and cells, merged ranges, conditional formatting and data validation ranges along with the formulas their rules and bounds are written as, hyperlinks, autofilter, frozen panes, row breaks, table extents, defined names, formula references on the sheet and on every other sheet pointing at it, the anchors a drawing places its shapes by, the cell each comment belongs to along with the box it appears in, the range a pivot table reads along with the block it is drawn in, and the range each chart series reads.

Refused, naming the thing: a formula we cannot rewrite with confidence.
A shape standing only on rows that are going away, since nothing is left to hang it from and dropping a chart in silence is worse than refusing.
Rows that take the whole of a pivot's source range with them, since a range attribute has no spelling for a reference pointing nowhere.
And a pivot built from consolidation ranges, whose own ranges are in a shape we do not read.

A drawing counts rows from zero where a sheet counts from one, and an absolute anchor names no row at all and does not move, which matches Excel.

Changed: drawings were refused in the first version, deliberately, so that something worked before the hardest part was attempted.
They move now.

Changed: comments were refused too, for a reason of their own.
A comment says which cell it belongs to in one part, and the box it appears in is positioned by a second written in VML, a legacy format we had no reader for, so moving the cell alone would have left the box behind.
Both parts move now, under one rule, and VML counts rows from zero the way a drawing does.
A comment whose cell went away is dropped, text and all, because that is what Excel does when a row carrying one is deleted.
That is the one place a comment parts company with a shape, which closes up against the rows that went rather than going with them.
The same VML part carries any form control and any header or footer image the sheet has, so those move on the same pass.

No library we test against reads a comment back, `exceljs` not even from its own file, so SheetJS is what proves the moved comment survives.
That Excel accepts a VML part we rewrote rather than copied was the open question, since the rewrite loses self-closing tag spelling the way every other rewritten part does.
Confirmed by hand in Excel on 2026-08-22.

Changed: pivot tables were refused too, and were the last thing on a sheet that was.
A pivot is three parts.
The cache definition says which range it reads, the cache records hold a copy of the rows in it, and the pivot table part says where the summary is drawn.
Moving the range is ordinary arithmetic, and so is moving the block when the pivot sits on the sheet that moved.
The copy of the rows is the interesting one, and we do not rewrite it.
It can be the largest part in the file, and rebuilding it means reading the data we just wrote back out, which is the kind of second pass the whole design avoids.
Instead the cache is marked `refreshOnLoad`, which ISO/IEC 29500 defines as the application refreshing the cache when the workbook is loaded.
That is the same bargain `fullCalcOnLoad` strikes for a stale formula result, and that one is confirmed working in Excel.

A cache whose source is written as a name rather than as a range needs nothing moved, since the name moves itself, but its copy of the rows is stale all the same, so it is marked for refresh and nothing else.
That case used to pass a save without being noticed at all, because the refusal only ever looked at a range.
So the blocker was catching the louder half of the problem and letting the quieter half through.

A cache reading an external query or a scenario is left alone.
Refreshing one of those on open could ask for credentials, which is a worse outcome than a stale figure, and a region cannot have moved its rows anyway.

Changed: chart series ranges were the last thing left pointing at the old cells, and they were never refused, only copied.
So a chart over a filled region was silently wrong, which is the failure this decision refuses everywhere else.
They move now.

A series is a formula in a `c:f` element, sheet qualified and absolute, so `shiftFormula` already knew how to move one and this is the same answer the rest of the write path gives rather than a new one.
Two files Excel wrote were read to confirm that shape before anything was built.
Unlike a pivot's source, a chart reference can carry `#REF!`, because it is a formula rather than a range attribute, so a series whose rows all went away is written the way Excel writes it rather than refusing the file.

Charts are found through the content types rather than by following relationships.
A chart hangs off a drawing, and a drawing off either a worksheet or a chartsheet, but the ranges a chart reads can name any sheet at all.
So walking out from the sheet that moved would miss a chart drawn anywhere else, and would miss one on a chart sheet entirely.
A chart is also the only part positioned by row that can read more than one sheet, so it is the only one given more than one move.

A chart carries its own copy of the values each series read, which we do not rewrite, and unlike a pivot cache it has no `refreshOnLoad` to ask for a rebuild with.
Excel re-reads the range on open regardless, so nothing has to be marked.
Confirmed by hand in Excel on 2026-08-22.

A cache built from consolidation ranges is refused.
It reads worksheet ranges of its own, spelled in a shape we do not read, so we cannot tell whether the moved rows are among them and cannot move them if they are.
That is the one pivot case where refusing is the answer rather than moving, and it costs a file whose ranges were all somewhere else.
That is the side to be wrong on, since the alternative is a total that quietly stops adding up.
A source type we have never seen is treated as consolidation, so an unfamiliar one refuses rather than being assumed harmless.

With that, `blockersFor` had nothing left to report and the module is gone.
What still stops a save is thrown where the work happens: a formula in `shift-formula`, a shape in `shift-drawing`, an extension list in `shift-sheet`, a pivot source in `shift-pivot`.
That is a better place for it, since the code that knows a thing cannot be moved is the code that tried.

A region shrinks to one row and no further.
A range whose every endpoint is deleted becomes `#REF!` in Excel, so one surviving row is what keeps a total written over the region alive, and it costs one blank formatted row on a run with no data at all.
Collapsing it entirely is a later option and changes nothing else.

Excel is more forgiving here than we assumed.
A range with one endpoint inside the deleted rows shrinks to what survives, rather than breaking, so `=SUM(C9:C13)` minus row 13 becomes `=SUM(C9:C12)`.
Only a reference with nothing left to point at dies, which is a single cell in a deleted row or a range wholly inside them.
Confirmed by hand in Excel on 2026-08-18.


#### What has to be known by the time each part is written

Almost every hard call in this decision follows from one constraint, so it is worth stating on its own rather than leaving it to be re-derived.

A file is written in one pass, front to back, and a part already written cannot be revised.
So a part can only depend on something learned before it was written.

The one thing that has to be learned is how far the rows moved, and it is learned by counting the caller's rows, which happens when the region is reached.
"Above" and "below" mean earlier and later in the sheet's rows, which is also earlier and later in the output, and that is the only reason the distinction matters.

```
        written                         knows the move?
  ---------------------------------------------------------
   1    the region's own sheet
          everything above the region    no
          the region's rows              being counted here
          everything below the region    yes
   2    every other sheet                yes
   3    the moved sheet's drawings,
        comments, VML and pivot parts,
        and every chart in the workbook  yes
   4    the workbook part, for its names yes
   5    a grown table's part             yes
   6    the styles part                  no move involved
```

Everything from the region's own rows onward already knows, so it needs nothing special.
Only the first box is a problem, and only when something in it reads rows below the region.
That is why the writer buffers what is above, checks whether any of it points down, and only counts the rows first when it has to.

Steps 2 to 5 are written after step 1 for this reason and no other.
The styles part is last for a reason of its own, which is that a date only learns which cell format it needs while its sheet is written.

The same constraint decides the multi-region rule.
Two regions on sheets whose formulas read each other would each need to be step 1, and there is no order in which both are.

A region's rows are counted as they are written, so nothing is held.

That is only possible because the writer owns the move.
The move used to be a transform sitting ahead of the writer, which meant it had to know how far the rows went before the writer had counted them, so every row was read first.
Now the writer renumbers as it goes, because only it knows when the region has been filled.

Two shapes still have to count first.
Whatever sits above the region goes out before its rows are counted, so if anything up there reads rows below the region it cannot be written until the count is known.
And two regions in one save, on sheets whose formulas read each other, cannot be ordered so that both learn about the other first.
Both fall back to counting first, which is what every region used to do.

This supersedes the table growth rule from decision 14.
A table with a totals row was refused because that row would have to move.
Now it moves, so a table grows whatever is under it, and the advice to put a total above a table goes with it.


## Open questions

- **Cell value typing (resolved).**
  A cell is a typed union: number, string, boolean, date, error, and formula.
  The caller reads `type` and gets a matching `value`, rather than raw strings to interpret.
  A formula cell carries its text as `value` and its cached result as `cachedValue`, and `ResolvedValue` is the union minus the formula variant, so a cached result cannot itself be a formula.
- **Scope of content for v1 (resolved).**
  Formulas, styles, number formats, and date interpretation are all in, not raw values only.
  A date comes back as a `Date`, from the serial number and its style in xlsx or the ISO value in ods.
  A formula reads as its cached value.
- **Shared strings and large sheets (resolved).**
  The sheet XML streams row by row and is never held.
  The shared strings table is read into memory and held, because cells reference it by index.
  So peak tracks the shared strings table plus a working window, not the sheet size, and a huge shared strings table is the remaining floor.
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
  The library ships one ES module built on web standards (`Uint8Array`, `ReadableStream`, `DecompressionStream`, `Blob`), so it runs in the browser as is, and the demo site reads a `File` there.
  Still open: running the library's own test suite in a real browser, not just Node, to confirm the stream and decompression code behaves the same.
- **String encoding when writing (resolved).**
  A written string cell is embedded inline (`inlineStr`), never added to the shared strings table.
  Shared strings dedupe and match what Excel itself emits, so files are smaller, but the writer would have to hold the whole table to dedupe against it, which fights the single streaming pass, and filling a template would mean rewriting a part we would otherwise copy through untouched.
  Inline strings need no global table, so a sheet writes in one pass and the shared strings part stays byte-identical to the source.
  The cost is a larger file when the same string repeats.
  Reading handles both, so nothing downstream cares.
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
  ODF keeps the whole spreadsheet in one `content.xml`, so listing sheets and reading a sheet each scan that part, unlike xlsx's per sheet parts.
  The list scan runs once at open.
- **`.ods` reading speed (profiled, deferred).**
  ODF keeps every sheet in one `content.xml` deflate stream, which cannot be seeked, so each access re-decompresses and re-parses from the start.
  On a 14 MB file that is 359 MB uncompressed: listing sheet names is about 3.2s (the XML reader tokenises cells it then discards, decompression alone is only about 0.5s), and reading a sheet re-parses from the start up to its table, up to 3.4s for the last sheet.
  Reading every sheet is about 25s, roughly one full parse per sheet.
  Profiled and left for now.
  A targeted byte scan for the `<table:table>` tags would cut the name list to about 0.5s, and skipping past earlier sheets would cut a single sheet read, but both trade the XML reader for hand rolled scanning, and reading every sheet at once is not a real use case.
  The memory story is unaffected and stays bounded.
  Revisit when a real workload hits it.
- **Website presents both formats (resolved).**
  The site covers `.ods` everywhere it matters now: the hero tagline, the page meta description, the demo copy, and a dedicated `.ods` block in the Benchmark section framed on the memory tradeoff (`ODS_BENCHMARK` in `website/app/site.ts`).
  The xlsx benchmark is labelled "Reading xlsx" so it reads as one of two scenarios, and the "Fast" feature blurb is scoped to xlsx so it does not contradict the slower `.ods` numbers.
- **Formula text is format-specific, and the read side leaks it.**
  The same formula is `SUM(A1:B1)` in xlsx and `of:=SUM([.A1:.B1])` in ods.
  That is a namespace prefix, a leading `=` and a different reference syntax, not a punctuation difference.
  Reading hands out each format's text as a plain string, so nothing stops a caller reading an ods formula and writing it into an xlsx as literal nonsense.
  The writer does not cause this, it makes it reachable.
  A format-neutral formula needs a parser that rewrites references and maps function names, which is a smaller job than the formula engine but still a real one.
  `formula()` exists as the boundary where that translation would go.
- **How faithful is a rewritten part.**
  Untouched entries are copied as bytes and are exactly identical, checksum and compressed size included, so nothing is recompressed.
  Writing into a table rewrites that table's part even when the table did not change size, because whether to copy an entry or rebuild it is decided before any row has been read and the size is only known afterwards.
  The result is semantically identical, so this costs byte-identity on one small part and nothing else.
  Five parts are rewritten on any edit: the edited sheets, `xl/styles.xml`, `xl/workbook.xml`, and, only when there was a calculation chain to drop, `[Content_Types].xml` and `xl/_rels/workbook.xml.rels`.
  Writing into a region adds the moved sheet's drawings, its comments, the VML positioning them and the pivot parts reading or drawn on it, because everything positioned by row has to move with the rows.
  A rewritten part is re-emitted from the XML event stream, which does not carry attribute order, self-closing tag spelling, comments or processing instructions.
  So it is semantically equivalent but not byte-identical, and the byte-identity test can only cover the parts we did not touch.
  Whether that gap ever matters is unknown.
- **Recalculation and the zip's shape (resolved).**
  Four things we could not prove with a test are no longer guesses, all confirmed on 2026-08-17 in Excel in the browser and in desktop Excel.
  An entry described after its data is accepted, which is what lets a sheet be deflated as it streams rather than buffered.
  A sheet with no `dimension` element is accepted, which a single pass cannot emit correctly because the element precedes the rows it describes.
  `fullCalcOnLoad` makes Excel recompute a cached formula result that our edit made stale.
  And dropping `xl/calcChain.xml` along with its content type override and its relationship leaves nothing dangling.
  That last one took two passes.
  The first template had no calculation chain, because `exceljs` writes none even for a workbook with formulas, so the code that removes one never ran.
  `scripts/manual-check.mjs` now splices one in so the branch is reachable.
  Desktop Excel is the stricter of the two and agrees.
  So the design decisions that only Excel could settle are settled, and what remains is coverage rather than doubt: the generated template has no charts or pivot tables, so the two fixtures saved from Excel and kept in `test/fixtures` cover those instead.
  `npm run manual-check` builds the file to try it with.
- **What a written file does not update.**
  A chart pointing at a fixed range does not extend to cover rows appended below it, since appending moves no rows for the range to follow.
  Filling a region does move them, and a chart's series follows.
  An Excel Table written by name does now grow, unless it has a totals row.
  One filled by row number through `writeRows` or `appendRows` still does not, because nothing in that call says the rows belong to it.
  A digitally signed workbook has its signature invalidated by any modification, which is inherent and not fixable.
  A template whose data region sits above a totals block is not served: writing past the region overwrites the totals, and we cannot detect it because we do not know those rows are a totals block.
  Decision 14 answers this for a template whose data region carries a name, since the name is the author telling us where the region ends.
  It stays true for one addressed only by coordinates.
- **Verifying the writer needs real files (partly done).**
  In place: the byte-identity test in `test/xlsx/write-fidelity.test.ts` fills a file `exceljs` wrote and asserts every entry we did not rewrite comes out with the same checksum, compressed size and bytes, including the theme and document properties, which we have no reader for at all.
  `exceljs` and `xlsx` both read the output back.
  The loop is guarded by naming parts it must have checked, so it cannot pass by checking nothing.
  `test/fixtures/pivot-template.xlsx` and `test/fixtures/chart-template.xlsx` are files Excel wrote, kept because `exceljs` can write neither part, and both are filled and asserted against in tests.
  Still missing: conditional formatting, which `exceljs` does write but not in a template that exercises the write path.
  Excel offering to repair a file is the failure that matters most and no library round-trip catches it.
  That is a manual check, listed in `MANUAL-CHECKS.md`.
- **Text outside the rows was moved by its tags alone (resolved).**
  A conditional format's rule and a data validation's bounds are formulas held as element text rather than as attributes, and everything outside a worksheet's rows was moved by rewriting attributes only.
  So the range a rule covered followed the rows while the rule itself kept naming the old ones, which highlights one row on the strength of another, and a validation went on bounding itself by rows that had moved.
  Both are the quiet kind of wrong this design refuses everywhere else, and both had been reachable since regions could move rows at all.
  What made it invisible is that the two paths that do move formula text, a cell's own and another sheet's, each track the tags they are inside, and the path for everything outside the rows had no such memory because until now nothing out there needed it.
  It has one now, in `insideFormula`, which the writer feeds each event.
  Only text a formula tag opened is moved, because the other text out there is a page header or footer, free text that can read like a reference and would be mangled by moving it.

- **Whether Excel plots a chart from the range or from its own copy of the values (resolved).**
  From the range.
  A chart part carries a cached copy of everything each series read, the way a pivot cache does, and we move the ranges and leave the copy alone.
  That is only good enough if Excel re-reads the range on open, and a pivot says so with `refreshOnLoad` while a chart has no equivalent to say it with, so this rested on behaviour rather than on anything in the file.
  Confirmed by hand in Excel on 2026-08-22: the fill writes eight rows under labels the cached copy does not hold, and all three charts draw them.
  So a chart needs nothing marked for refresh, which is where this differs from a pivot.

- **Charts in the newer chartex format (open).**
  Treemap, sunburst, histogram, box and whisker, waterfall, funnel and map charts are written to a different part, with its own content type, which `readChartPaths` does not look for.
  So one of those over a filled region is silently wrong, which is the thing this design refuses.
  Its references look like they are spelled the same way, in an element also named `f`, but that is not in ECMA-376 and we have not read a real one, so building on it would be guessing.
  Settling it needs a file with a waterfall chart in it, the same missing-file problem the pivot and chart work both had.
  Until then the choice is between refusing a save when the package holds one and leaving it stale, and that is not decided.

- **A chartsheet is listed as a worksheet (resolved by leaving it out).**
  It used to be, because `readWorkbook` took every `<sheet>` in the workbook part, so a chart sheet came back from `worksheets` as though it held cells.
  `worksheets` now means the sheets that hold rows, which is the same split Excel's own object model makes between its `Sheets` and its `Worksheets`, so a chart sheet is not among them and asking for one by name says it was not found.
  The alternative was to keep it and mark it, which we turned down because `worksheets` would then no longer say what you can read and every caller would carry the check.
  The cost is that a caller addressing a sheet by position sees the positions after a chart sheet move up by one.
  Which kind a sheet is shows nowhere on the `<sheet>` element, only in the type of the relationship it points through, so that is what decides it, and only a type that positively names something other than a worksheet drops one.
  A relationship that is missing or names no type has said nothing, and losing a sheet we cannot place would be worse than failing when its rows are asked for.
  The whole sheet order is still read, because two things count every sheet and not only the worksheets: a defined name's scope, which is a position in that order, and the sheet ids, since an added worksheet must not take one a chart sheet already holds.
  That is why `WorkbookInfo` carries `highestSheetId` rather than letting the editor take the largest id it can see.

- **Reading by name (parked).**
  The writer parses defined names anyway, so exposing the resolved names and their extents to a reader is a small surface for free parsing.
  The case that would earn it is an input workbook, a filled-in form whose values live in named cells, where the name survives an author moving the cell and a coordinate does not.
  Not a priority.
  A region read is a different shape from the row stream the read API is, and nobody has asked for either.

- **What a region costs (resolved).**
  A million rows into a region finishes under a 150MB heap, the same as `appendRows`, where it once needed 500MB.
  Its rows are counted as they are written rather than beforehand, so nothing is held in proportion to the data.
  Getting there meant moving the row shifting from a transform ahead of the writer into the writer itself.
  Four attempts to shrink the holding instead had moved the time from 7.4s to 5.6s and the memory not at all, which was the sign that the holding was the design rather than an inefficiency.
  The write benchmark has `region-stream` and `region-load` modes, and `--cap` is the measure worth quoting, since peak RSS counts memory the runtime has not given back.
  What is left is a flat 15MB or so above `appendRows`, the same at fifty thousand rows as at a million, so it is a constant and not something that grows.
  Two shapes still count first, both recorded in decision 15: something above the region reading rows below it, and two regions whose sheets read each other.

- **Whether Excel really rebuilds a pivot cache on open (resolved).**
  We mark a moved cache `refreshOnLoad` and leave its copy of the rows as it was, so between opening the file and the refresh happening the cache describes data that is no longer there.
  Confirmed by hand in Excel on 2026-08-22.
  A pivot over eight written rows drew a region that appears in none of the cached rows, and totalled the data we wrote rather than the data the file was saved with.
  No prompt appeared first, so this is a silent guarantee rather than something a caller has to warn anyone about, and a `recordCount` that no longer matches is not a repair prompt either.
  What a reader that ignores `refreshOnLoad` shows is still unknown, and is the caveat the README carries.
  We considered also setting `invalid="1"`, which the spec defines as the cache needing a refresh, and did not, because a cache marked invalid whose refresh the user declines may show nothing at all, where a stale one shows an old figure.
  With `refreshOnLoad` confirmed on its own there is no reason to revisit that.

- **A pivot drawn on the sheet whose rows moved.**
  The pivot table part's `location` is moved as a plain range, and the counts inside it are relative to that range's top left, so they are left alone.
  This is unit tested and has never been in front of Excel, because the template we have puts the pivot on its own sheet, which is where Excel puts one by default.
  The case is real, a pivot below a data region on one sheet, and worth covering the next time a template is made.

- **A consolidation pivot source (resolved by refusing).**
  A cache reading an external query or a scenario is left alone, on the reasoning that a region cannot have moved its rows.
  That does not hold for a consolidation source, which reads worksheet ranges of its own.
  It briefly went stale quietly, which is the failure this whole design exists to avoid, so it refuses now.
  Refusing needs only the source type, which the spec spells out, where moving it would need the shape of the ranges inside, which we have no file to learn from.
  So this is settled in the sense that nothing is quietly wrong, and open in the sense that a template with one cannot be filled at all.
  Moving them is the work to do if anyone hits it, and it starts with a file that has one.

- **A grown table's `sortState`.**
  A table part can hold a `sortState` recording the last sort applied to it, with its own `ref` over the data.
  We rewrite the table's extent and its autofilter range when it grows, and leave `sortState` as it was.
  Whether a stale one matters is unknown.
  It records what was done rather than describing the table, so the likely answer is that Excel ignores it until someone sorts again, but that is reasoning rather than evidence, and growing it wrongly would be worse than leaving it.
  `exceljs` writes none, so no fixture here has one to look at.

- **Writing `.ods`.**
  Deferred, and not decided against.
  Writing only one of the two formats we read is a real asymmetry and we are not calling it the end state.
  What it costs is now clearer than it was, and the cost went up rather than down.
  It shares the zip and XML writing layers and nothing else, since the fidelity work is per format.
  It forces the formula translation question above, because `formula()` cannot pass its text through untranslated.
  And the copy-through guarantee is structurally weaker there.
  In xlsx an edit rewrites the sheet part plus at most four small ones and every other entry is copied byte-identical, which `test/xlsx/write-fidelity.test.ts` asserts.
  ODF keeps every sheet, every automatic style and every conditional format in one `content.xml`, so editing one cell sends the whole document back through the XML event stream, which as noted above does not carry attribute order, self-closing spelling, comments or processing instructions.
  The gap that covers a handful of parts in xlsx would cover everything in ods.
  Decision 13 removes one cost, since there is no styling API to build twice.
  Decision 14's anchors port, since ODF has named expressions with cell range addresses, spelled `Sheet1.B4:Sheet1.D20` rather than `Sheet1!$B$4:$D$20`.
  Revisit once the anchor work has landed for xlsx and we know what the second implementation would actually repeat.
