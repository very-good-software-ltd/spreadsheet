# Manual checks

Things no test in this repo can prove, to be run by hand before a release that
changes the write path.

The reason these exist at all: `exceljs` and `xlsx` reading our output back
proves the file is well formed to another JavaScript library. It does not prove
Excel accepts it. Excel offering to repair a file is the failure that matters
most to a user and the one nothing in CI will catch.

Use real Excel, not LibreOffice. LibreOffice is more forgiving, so passing there
proves less than it looks.

Desktop Excel and Excel in the browser are different implementations, and the
desktop one is stricter. Passing in the browser is real evidence but not the whole
check, so run both if you can.


## What has been checked so far

Checked on 2026-08-17, in Excel in the browser and in desktop Excel. No repair
prompt, every expectation on the Checks sheet held, and document properties came
through with the author intact.

That settles four things no test here can reach. An entry described after its
data is accepted. A sheet with no `dimension` element is accepted. A stale cached
formula result is recalculated. And removing the calculation chain leaves no
dangling content type override or relationship behind.

So the parts of the design that could only be confirmed by Excel are confirmed.

Checked again on 2026-08-18, after named regions landed. The `Summary` sheet
holds: the row written by name is in the right place, the two rows of the region
we were not given come back empty with their formatting, and the labels either
side of it are untouched. So clearing works in Excel and stops at the region's
edges.

What is still worth a pass before a release is a real template, since the
generated one has no charts or pivot tables in it. The `Movements` entry in
Formulas, Name Manager has not been looked at yet, and it is the one item on the
Checks sheet that is not visible from the grid.


## Start here

```sh
npm run manual-check
```

That writes two files into `manual-check/`:

- `template.xlsx`, a made-up corporate template with a merged heading, column
  widths, frozen panes, conditional formatting, a pre-formatted data region, a
  named region on its `Summary` sheet, a total formula and a calculation chain.
- `filled.xlsx`, that template after this library filled it in.

Open `filled.xlsx` in Excel. It has a **Checks** sheet listing every cell to look
at and what you should see there, so the file carries its own checklist and you
do not have to read this document while you work.

Everything is fine if the file opens with no prompt and the Checks sheet's
expectations all hold. Work down that sheet, then come back here for the two
things it cannot cover.


## What the generated template cannot cover

`exceljs` writes the template, and it cannot write charts or pivot tables. Those
are the parts most likely to be lost by a writer that rebuilds a file from a
model, so they are the most valuable thing to check and the one thing the
generated file misses.

To cover them, put a real template at `manual-check/template.xlsx` and run the
script again. It uses your file instead of generating one. Then check:

- [ ] Charts still render, and still point at their data.
- [ ] Pivot tables still open and still refresh.
- [ ] Data validation still restricts what you can type.
- [ ] Print setup, headers and footers survive.
- [ ] Defined names still resolve.
- [ ] Any macros still run.

The script writes into `Report` by name and into rows 9 onward, and into a region
named `Movements` on a sheet called `Summary`, so a template built differently
needs those lines in `scripts/manual-check.mjs` adjusted to match. If your
template names a region of its own, point the `writeRegion` call at that name
instead, since a region the author drew in Excel is closer to the real case than
one we generated.

The script keeps whatever is already at `manual-check/template.xlsx` and only
generates one when nothing is there. So after a change to the generated template,
delete both files in `manual-check/` before running it, or you will be filling
yesterday's file.


## Why each check is there

Background for anything on the Checks sheet that looks arbitrary.

**Opening with no repair prompt** is the one that covers the most ground. An
entry we compress ourselves cannot know its checksum or compressed size until
the data is written, and the local header comes first, so the entry is described
afterwards in a data descriptor with bit 3 of the general purpose flag set (see
`src/zip/native-zip-writer.ts`). Buffering the entry instead would remove the
need for this, at the cost of holding a whole sheet in memory, which is the thing
the library promises not to do. So if Excel rejects it, the fix is a real design
change and not a small one. Apache POI's streaming writer produces xlsx the same
way, which is why we expect it to pass, but expecting is not knowing.

**The total reading 1400 rather than 0** is recalculation. The template caches a
result of 0, which was right when it was saved and is wrong for the data we
wrote. We drop `xl/calcChain.xml` and set `fullCalcOnLoad` in the workbook part
to force Excel to work it out again.

Dropping that part also means removing its content type override and its
relationship, and a dangling one of either is exactly what makes Excel offer to
repair a file. `exceljs` writes no calculation chain even for a workbook with
formulas, so the script splices one into the template on purpose. Without it that
whole path would never run and the check would prove nothing about it.

**Rows 11 and 12 looking like rows 9 and 10** is `inheritFrom`. The template's
formatted data region is only two rows deep and we wrote four rows into it, so
the last two had no formatting to keep and copied row 9's instead.

**C4 showing a date and staying bold** is the styles rewrite. That cell had a
bold font and no number format, so writing a date had to copy its cell format and
add a date format to the copy. Replacing the format instead would have shown the
date correctly and lost the bold. **C5** is the other half: it already had a date
format, so we left it alone rather than adding a second one.

**Nothing checks the `dimension` element directly**, because its absence has no
visible symptom. The template has one and the output does not, so opening the file
at all is the check. It records the used range and sits before the row data, so a
single streaming pass cannot know the final extent by the time it would have to
be written. It is optional in the schema, so we leave it out. If scrolling to the
end of a written sheet behaves oddly, or Excel's used-range shortcuts land in the
wrong place, that is where to look.


## Why the named region is checked twice

The `Summary` sheet checks the two halves of what a named region promises.

That it went in at all covers the write finding its place from the name rather
than from a row number, which is the whole reason the feature exists.

That rows 4 and 5 come back empty covers the harder half. Those cells held 999
in the template, formatted exactly like the row we did fill. If clearing failed,
the sheet would look completely normal and the totals underneath would be wrong,
which is the failure this library exists to avoid rather than cause. The labels
in column A and the text in column E are there to prove the clearing stopped at
the region's edges instead of taking the whole row.

Name Manager still listing `Movements` matters because `xl/workbook.xml` is one
of the few parts we rewrite rather than copy. A name lost in that rewrite would
not show up in the filled file at all, and the next run against it would fail to
find the region. A test covers that too, by filling the output a second time.


## Whether a name's case matters

We look a defined name up without regard to case, so a caller asking for `data`
finds a name the template spells `Data`. That is what we believe Excel does, and
no test here can prove it, because a file cannot show you a rule about names it
does not contain.

To check it, in Excel: Formulas, Name Manager, New, and define `Data` over any
range. Then New again and try to define `data` over a different range.

We expect Excel to refuse the second one, saying the name already exists. If it
accepts both, then a workbook can hold two names differing only in case, our
lookup would pick whichever comes first in the file, and the rule needs changing
to match exactly before falling back.


## Known and accepted

Not failures, so do not chase them:

- An Excel Table does not grow to cover appended rows.
- A chart pointing at a fixed range does not extend to cover them either.
- A digital signature is invalidated. Any modification does that, and nothing can
  be done about it.
- A rewritten part is not byte-identical to its source, only equivalent. Attribute
  order, self-closing tag spelling and comments are not preserved. Only the parts
  we copy are identical, and a test already asserts that.
