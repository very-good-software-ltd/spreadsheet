# Manual checks

Things no test in this repo can prove, to be run by hand before a release that
changes the write path.

The reason these exist at all: `exceljs` and `xlsx` reading our output back
proves the file is well formed to another JavaScript library. It does not prove
Excel accepts it. Excel offering to repair a file is the failure that matters
most to a user and the one nothing in CI will catch.

Use real Excel, not LibreOffice. LibreOffice is more forgiving, so passing there
proves less than it looks.


## Start here

```sh
npm run manual-check
```

That writes two files into `manual-check/`:

- `template.xlsx`, a made-up corporate template with a merged heading, column
  widths, frozen panes, conditional formatting, a pre-formatted data region and a
  total formula.
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

The script writes into `Report` by name and into rows 9 onward, so a template
whose first sheet is called something else needs those two lines in
`scripts/manual-check.mjs` adjusted to match.


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
to force Excel to work it out again. Dropping that part also means removing its
content type override and its relationship, so this check covers whether we left
anything dangling.

**Rows 11 and 12 looking like rows 9 and 10** is `inheritFrom`. The template's
formatted data region is only two rows deep and we wrote four rows into it, so
the last two had no formatting to keep and copied row 9's instead.

**C4 showing a date and staying bold** is the styles rewrite. That cell had a
bold font and no number format, so writing a date had to copy its cell format and
add a date format to the copy. Replacing the format instead would have shown the
date correctly and lost the bold. **C5** is the other half: it already had a date
format, so we left it alone rather than adding a second one.

**Nothing checks the `dimension` element directly**, because its absence has no
visible symptom. It records the used range and sits before the row data, so a
single streaming pass cannot know the final extent by the time it would have to
be written. It is optional in the schema, so we leave it out. If scrolling to the
end of a written sheet behaves oddly, or Excel's used-range shortcuts land in the
wrong place, that is where to look.


## Known and accepted

Not failures, so do not chase them:

- An Excel Table does not grow to cover appended rows.
- A chart pointing at a fixed range does not extend to cover them either.
- A digital signature is invalidated. Any modification does that, and nothing can
  be done about it.
- A rewritten part is not byte-identical to its source, only equivalent. Attribute
  order, self-closing tag spelling and comments are not preserved. Only the parts
  we copy are identical, and a test already asserts that.
