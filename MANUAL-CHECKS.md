# Manual checks

Things no test in this repo can prove, to be run by hand before a release that
changes the write path.

The reason these are here at all: `exceljs` and `xlsx` reading our output back
proves the file is well formed to another JavaScript library. It does not prove
Excel accepts it. Excel offering to repair a file is the failure that matters
most to a user and the one nothing in CI will catch.

Run these in real Excel, not LibreOffice. LibreOffice is more forgiving and
passing there proves less than it looks.


## Writing

Produce a file with the write path, open it in Excel, and check each of these.

- [ ] **The file opens with no repair prompt.**
      This is the one that covers the most ground. If Excel offers to repair the
      file, stop and find out why before checking anything else.

- [ ] **Entries described after their data are accepted.**
      An entry we compress ourselves cannot know its checksum or compressed size
      until the data is written, and the local header comes first, so the entry
      is described afterwards in a data descriptor with bit 3 of the general
      purpose flag set (see `src/zip/native-zip-writer.ts`).
      Buffering the entry instead would remove the need for this, at the cost of
      holding a whole sheet in memory, which is the thing the library promises
      not to do. So if Excel rejects this, the fix is a real design change and
      not a small one.
      Apache POI's streaming writer produces xlsx the same way, which is why we
      expect this to pass, but expecting is not knowing.

- [ ] **Formulas recalculate on open.**
      Change a value that an existing formula depends on, then open the file.
      The formula's result should reflect the new value, not the result cached
      when the template was saved.
      We drop `xl/calcChain.xml` and set `fullCalcOnLoad` in the workbook part
      to force this.

- [ ] **Dropping `xl/calcChain.xml` leaves nothing dangling.**
      The part is removed, so its content type override and its relationship
      have to go too, or Excel has to tolerate them pointing at nothing.
      We remove both. Confirm a file that originally had a calc chain still
      opens clean.

- [ ] **Omitting `dimension` is accepted.**
      The element records the used range and sits before the row data, so a
      single streaming pass cannot know the final extent by the time it would
      have to be written. It is optional in the schema, so we leave it out.
      Confirm Excel computes the range itself, and that scrolling to the end of
      a written sheet behaves normally.

- [ ] **A date written into an unformatted cell renders as a date.**
      Not as a five digit number. This exercises the number format we add to
      `xl/styles.xml`.

- [ ] **A date written into a cell that already had corporate formatting keeps it.**
      Font, fill and borders intact, and now rendering as a date. This exercises
      cloning the cell's existing format rather than replacing it.


## Template fidelity

Fill a real template that contains all of these, then open the result.

- [ ] Charts still render, and still point at their data.
- [ ] Pivot tables still open and still refresh.
- [ ] Merged cells, conditional formatting and data validation survive.
- [ ] Frozen panes, column widths and row heights survive.
- [ ] Print setup and headers or footers survive.
- [ ] Defined names still resolve.

Known and accepted, so not failures:

- An Excel Table does not grow to cover appended rows.
- A chart pointing at a fixed range does not extend to cover them either.
- A digital signature is invalidated. Any modification does that, and nothing
  can be done about it.
