# Manual checks

Things no test in this repo can prove, to be run by hand before a release that changes the write path.

The reason these exist at all: `exceljs` and `xlsx` reading our output back proves the file is well formed to another JavaScript library.
It does not prove Excel accepts it.
Excel offering to repair a file is the failure that matters most to a user and the one nothing in CI will catch.

Use real Excel, not LibreOffice.
LibreOffice is more forgiving, so passing there proves less than it looks.

Desktop Excel and Excel in the browser are different implementations, and the desktop one is stricter.
Passing in the browser is real evidence but not the whole check, so run both if you can.


## What has been checked so far

Checked on 2026-08-17, in Excel in the browser and in desktop Excel.
No repair prompt, every expectation on the Checks sheet held, and document properties came through with the author intact.

That settles four things no test here can reach.
An entry described after its data is accepted.
A sheet with no `dimension` element is accepted.
A stale cached formula result is recalculated.
And removing the calculation chain leaves no dangling content type override or relationship behind.

So the parts of the design that could only be confirmed by Excel are confirmed.

Checked again on 2026-08-18, after named regions landed.
The `Summary` sheet holds: the row written by name is in the right place, the two rows of the region we were not given come back empty with their formatting, and the labels either side of it are untouched.
So clearing works in Excel and stops at the region's edges.

Name Manager still lists `Movements` over its range, so the names in `xl/workbook.xml` survive the rewrite of that part.
Every expectation on the Checks sheet now holds.

Checked on 2026-08-19, after the writer took ownership of moving rows and after `Ledger` gained the rows below its table.
Every expectation on the Checks sheet holds.

That covers both directions of the model on one run.
`Summary` shrinks a region, pulls what is below it up and rewrites `SUM(B3:B5)` to `SUM(B3:B3)`.
`Ledger` grows one, pushes what is below it down and stretches `SUM(D3:D4)` to `SUM(D3:D7)`.
A total naming the table rather than its rows stays untouched and still adds up.

Checked on 2026-08-18, after drawings could move.
The picture on `Summary` starts on row 7 with one empty row between it and `Total` at row 5.
The template has `Total` at row 7 and the picture at row 9, also one row apart, so the anchor came up by the same two rows the region lost and Excel renders it where we put it.

Checked on 2026-08-18, after rows could move.
`Summary!B5` reads 120 with `SUM(B3:B3)` in the formula bar, where the template has `SUM(B3:B5)` at `B7`.
February and March are nowhere on the sheet.
`Ledger!D1` still reads 150.

That settles the largest thing the writer does.
The rows went, everything under them came up, and the formula was rewritten to match rather than left pointing at rows that are no longer there.
Both figures being right also means Excel opened the file and recalculated it, since a cached result would have been wrong for the data we wrote.

Checked on 2026-08-18, after tables landed.
`Ledger!D1` reads 150, so `SUM(Entries[Amount])` covered all five rows written into a table that had room for two.
A structured reference resolves over the table's extent and nothing else, so that number is Excel agreeing the table grew, not just that the rows are on the sheet.
Growing a table's extent and its filter's is therefore confirmed.

Checked on 2026-08-22, after comments could move.
The file opens with no repair prompt and `Summary` shows the two comments that should be there, at `A1` where it stayed and at `A5` where the template had it at `A7`.
The one on the row the region took out is gone.

So a VML part we rewrote rather than copied is accepted, which was the open question.
Excel does not mind losing self-closing tag spelling there any more than it does anywhere else, and the two parts a comment is written across came out agreeing with each other.

What is still worth a pass before a release is a real template, since the generated one has no charts or pivot tables in it.


## Start here

```sh
npm run manual-check
```

That writes two files into `manual-check/`:

- `template.xlsx`, a made-up corporate template with a merged heading, column widths, frozen panes, conditional formatting, a pre-formatted data region, a named region on its `Summary` sheet, an Excel Table on its `Ledger` sheet, a total formula and a calculation chain.
- `filled.xlsx`, that template after this library filled it in.

Open `filled.xlsx` in Excel.
It has a **Checks** sheet listing every cell to look at and what you should see there, so the file carries its own checklist and you do not have to read this document while you work.

Everything is fine if the file opens with no prompt and the Checks sheet's expectations all hold.
Work down that sheet, then come back here for the two things it cannot cover.


## What the generated template cannot cover

`exceljs` writes the template, and it cannot write charts or pivot tables.
Those are the parts most likely to be lost by a writer that rebuilds a file from a model, so they are the most valuable thing to check and the one thing the generated file misses.

To cover them, put a real template at `manual-check/template.xlsx` and run the script again.
It uses your file instead of generating one.
Then check:

- [ ] Charts still render, and still point at their data.
- [ ] Pivot tables still open and still refresh.
- [ ] Data validation still restricts what you can type.
- [ ] Print setup, headers and footers survive.
- [ ] Defined names still resolve.
- [ ] Any macros still run.

The script writes into `Report` by name and into rows 9 onward, and into a region named `Movements` on a sheet called `Summary`, so a template built differently needs those lines in `scripts/manual-check.mjs` adjusted to match.
If your template names a region of its own, point the `writeRegion` call at that name instead, since a region the author drew in Excel is closer to the real case than one we generated.

The script keeps whatever is already at `manual-check/template.xlsx` and only generates one when nothing is there.
So after a change to the generated template, delete both files in `manual-check/` before running it, or you will be filling yesterday's file.


## Why each check is there

Background for anything on the Checks sheet that looks arbitrary.

**Opening with no repair prompt** is the one that covers the most ground.
An entry we compress ourselves cannot know its checksum or compressed size until the data is written, and the local header comes first, so the entry is described afterwards in a data descriptor with bit 3 of the general purpose flag set (see `src/zip/native-zip-writer.ts`).
Buffering the entry instead would remove the need for this, at the cost of holding a whole sheet in memory, which is the thing the library promises not to do.
So if Excel rejects it, the fix is a real design change and not a small one.
Apache POI's streaming writer produces xlsx the same way, which is why we expect it to pass, but expecting is not knowing.

**The total reading 1400 rather than 0** is recalculation.
The template caches a result of 0, which was right when it was saved and is wrong for the data we wrote.
We drop `xl/calcChain.xml` and set `fullCalcOnLoad` in the workbook part to force Excel to work it out again.

Dropping that part also means removing its content type override and its relationship, and a dangling one of either is exactly what makes Excel offer to repair a file.
`exceljs` writes no calculation chain even for a workbook with formulas, so the script splices one into the template on purpose.
Without it that whole path would never run and the check would prove nothing about it.

**Rows 11 and 12 looking like rows 9 and 10** is `inheritFrom`.
The template's formatted data region is only two rows deep and we wrote four rows into it, so the last two had no formatting to keep and copied row 9's instead.

**C4 showing a date and staying bold** is the styles rewrite.
That cell had a bold font and no number format, so writing a date had to copy its cell format and add a date format to the copy.
Replacing the format instead would have shown the date correctly and lost the bold. **C5** is the other half: it already had a date format, so we left it alone rather than adding a second one.

**Nothing checks the `dimension` element directly**, because its absence has no visible symptom.
The template has one and the output does not, so opening the file at all is the check.
It records the used range and sits before the row data, so a single streaming pass cannot know the final extent by the time it would have to be written.
It is optional in the schema, so we leave it out.
If scrolling to the end of a written sheet behaves oddly, or Excel's used-range shortcuts land in the wrong place, that is where to look.


## Why the Summary sheet is the important one

It is the only check that covers rows actually moving, which is the largest thing the writer does and the one with the most ways to be silently wrong.

The region covers rows 3 to 5 and is given one row, so two rows have to go.
What to look at, in order of how badly it fails:

`B5` should read 120 and its formula bar should show `SUM(B3:B3)`.
In the template that formula is `SUM(B3:B5)` sitting at `B7`.
So this one cell proves three things at once: the rows went, everything under them came up, and the formula was rewritten to match rather than left pointing at rows that no longer exist.
A wrong number here is the failure, not a missing one.

February and March should be nowhere on the sheet.
Excel deletes a whole row, so labels beside the region go with it.
That is the deliberate cost of whole-row moves and it is worth seeing rather than reading about.

Name Manager should still list `Movements`, now over `Summary!$B$3:$D$3` rather than `$B$3:$D$5`.
If the name did not shrink with the region, filling the output a second time would write into rows that are no longer the data.
A test covers the same ground by filling twice, but only Excel can confirm it agrees the name is still well formed.


## Why the Ledger sheet has totals above and below

The `Ledger` sheet has a table with room for two rows, and the fill writes five, so the table has to grow by three.

`D1` holds `SUM(Entries[Amount])`, above the table rather than below it, and that cell is the point of the whole check.
A structured reference names the table instead of a range of rows, so if the table really grew, the total covers all five entries and reads 150.
If the extent did not move, Excel treats the extra rows as ordinary cells outside the table and the total reads 60, which is the first two rows we overwrote plus nothing else.
A wrong number here is the failure, not a missing one.

`D1` is above the table and `D9` is below it, and they check different things.
`D1` names the table, so it keeps covering everything without being touched.
`D9` is written over the table's rows as a range, `SUM(D3:D4)` in the template, so it has to be rewritten to `SUM(D3:D7)` and moved down three rows along with the `Checked` label beside it and the line of plain text under that.

That is the half of the model the `Summary` sheet cannot show.
`Summary` shrinks a region and pulls what is below it up.
`Ledger` grows one and pushes what is below it down.
Between them the two directions are covered, and they have to be on separate sheets because two regions cannot share one: writing either would move the rows the other was aimed at.

Clicking inside the last row and looking for Table Design on the ribbon asks the same question more weakly, and is only worth doing if the total looks wrong and you want to see where it went wrong.


## Why there is a picture on the Summary sheet

It stands in for a chart.
`exceljs` cannot write one, and a chart is anchored to rows exactly the way a picture is, so this covers the same mechanism with the tools available.

It is anchored below the Total, which is itself below the region, with one empty row between them.
Two rows come out of the region, so both have to come up by two and that single empty row has to stay a single empty row.
Three empty rows means the picture did not move.

This is the check most worth repeating against a real template, since a real one will have a chart rather than a picture, and a chart also holds its own series ranges, which we rewrite as formulas rather than as anchors.


## Why the Summary sheet carries comments

This is the one part of the file no library can check for us.
`exceljs` reads no comment back, not even from a file it wrote itself, so nothing in CI can prove a comment survived a save.
SheetJS does read them, and a test uses it to prove the moved comment is still there under the right cell, but that only proves another JavaScript library is happy.

There are three of them on `Summary`, one above the region, one on a row that goes away and one below it, so a single run covers moving, staying and being dropped.

The part positioning the boxes is VML, a legacy format, and we rewrite it rather than copy it.
That rewrite loses self-closing tag spelling the way every other rewritten part does, and whether Excel minds is exactly what this check answers.
It does not, as of 2026-08-22, so what is left is a regression check rather than an open question.
A comment appearing on the wrong row, or a marker with no box behind it, means the two parts disagree.


## Whether a name's case matters

Settled on 2026-08-18, so this needs no repeating unless the lookup changes.

We look a defined name up without regard to case, so a caller asking for `data` finds a name the template spells `Data`.
No test here can prove that is Excel's rule, because a file cannot show you a rule about names it does not contain.

Checked by hand instead: Formulas, Name Manager, New, define `Data` over a range, then New again and define `data` over a different one.
Excel refuses the second with "This name already exists.
Names should be unique".
So a workbook cannot hold two names differing only in case, and matching without regard to case cannot pick the wrong one of a pair that cannot exist.


## Known and accepted

Not failures, so do not chase them:

- An Excel Table does not grow to cover appended rows.
- A chart pointing at a fixed range does not extend to cover them either.
- A digital signature is invalidated.
  Any modification does that, and nothing can be done about it.
- A rewritten part is not byte-identical to its source, only equivalent.
  Attribute order, self-closing tag spelling and comments are not preserved.
  Only the parts we copy are identical, and a test already asserts that.
