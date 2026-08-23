// Produces the pair of files the manual checks in MANUAL-CHECKS.md need: a
// template, and that template filled in by this library. Open the filled one in
// Excel and work down its "Checks" sheet, which lists what to look at and what
// you should see.
//
// Drop your own file at manual-check/template.xlsx and it is used instead of the
// generated one.
//
// It also fills the two templates kept with the tests, giving three files to open
// on the same run, since neither a pivot table nor a chart can be generated.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { date, Workbook } from "../dist/index.js";

const OUTPUT_DIR = "manual-check";
const TEMPLATE = join(OUTPUT_DIR, "template.xlsx");
const FILLED = join(OUTPUT_DIR, "filled.xlsx");

// A pivot table cannot be generated, since exceljs cannot write one, so the pivot
// check runs off a file saved from Excel and kept with the tests.
const PIVOT_TEMPLATE = join("test", "fixtures", "pivot-template.xlsx");
const PIVOT_FILLED = join(OUTPUT_DIR, "pivot-filled.xlsx");

// A chart cannot be generated either, for the same reason, so this one is also a
// file saved from Excel and kept with the tests.
const CHART_TEMPLATE = join("test", "fixtures", "chart-template.xlsx");
const CHART_FILLED = join(OUTPUT_DIR, "chart-filled.xlsx");

const MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SPREADSHEET_TYPES = "application/vnd.openxmlformats-officedocument.spreadsheetml";
const CALC_CHAIN_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E7FF" } };
const THIN_BORDER = { style: "thin", color: { argb: "FF999999" } };
const BORDERED = { top: THIN_BORDER, left: THIN_BORDER, bottom: THIN_BORDER, right: THIN_BORDER };

// What the filled file should look like once Excel has opened it. Written into
// the output as its own sheet, so the file carries its own checklist.
//
// Do not write an expectation that is only a number. These are string cells, and
// Excel puts a "number stored as text" warning on any text cell whose content is a
// valid number, which reads as a fault in the file rather than in the checklist.
// A one pixel PNG, scaled up where it is placed. The picture does not matter, only
// where it ends up.
const PIXEL_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const EXPECTATIONS = [
  ["Where", "What you should see", "Why it matters"],
  ["The file itself", "Opens with no repair prompt", "An entry described after its data is accepted"],
  ["Report!C3", "Northwind Traders", "A plain value was written"],
  ["Report!C4", "2026-03-01, and still bold", "A date cloned the cell's format instead of replacing it"],
  ["Report!C5", "15/03/2026", "A cell already formatted as a date kept its own format"],
  ["Report!C13", "1400, not 0", "Every formula recalculated on open"],
  [
    "Report rows 9 and 10",
    "Filled in, borders and 1,234.00 amounts unchanged",
    "Rows written over the template's formatted region kept it",
  ],
  [
    "Report rows 11 and 12",
    "Bordered and formatted the same as rows 9 and 10",
    "Rows past the formatted region inherited its formatting",
  ],
  ["Report!A1:C1", "One merged blue heading", "A merged range survived"],
  ["Report column A", "Wider than the others", "Column widths survived"],
  ["Report row 8", "Still frozen when you scroll", "Frozen panes survived"],
  ["Report!B9:B12", "Red where the quantity is over 10", "Conditional formatting survived"],
  ["Notes!A1", "Left untouched", "A sheet we never opened is byte-identical"],
  [
    "Sheet tabs",
    "Report, Notes, Summary, Ledger, Checks",
    "A worksheet was added, with its relationship and content type",
  ],
  [
    "Summary row 3",
    "January, 120, 45, 165, and still bordered",
    "A region addressed by the name its author gave it was written, not by row number",
  ],
  [
    "Summary rows 4 and 5",
    "Gone. February and March are not on the sheet at all",
    "A region given one row keeps one row, and the two it did not need were taken out whole",
  ],
  [
    "Summary!A5",
    "Reads Total, where the template had it at A7",
    "Everything under the region came up by the two rows that went away",
  ],
  [
    "Summary!B5",
    "Reads 120, and the formula bar shows SUM(B3:B3)",
    "The total was written over the region's rows and followed it as it shrank, instead of breaking or summing the wrong ones",
  ],
  ["Summary!E3", "Still says 'keep me'", "A row that stayed kept everything on it"],
  [
    "Summary!A5",
    "Total is red, not black",
    "A conditional format below the region came up with the rows and its rule came with it. Black means the range moved and the rule stayed behind, reading a row that is now empty",
  ],
  [
    "Summary!C5, the bordered cell left of the arrow on row 5",
    "Type 50, accepted. Type 500, Excel refuses it",
    "A data validation's bounds are formulas too, and its upper one is the total beside it. Refusing 50 as well means the bound stayed behind on a cell that is now empty",
  ],
  [
    "The image on Summary",
    "Starts on row 7, with exactly one empty row between it and Total",
    "A picture anchored below the region came up with the rows. Three empty rows means it did not move, and a chart is anchored the same way",
  ],
  [
    "Summary!A1 and Summary!A5",
    "Both show a comment marker, reading 'Stays where it is' and 'Comes up with the total'",
    "A comment below the region came up with the rows, and both halves of it came with it, the cell it belongs to and the box it appears in",
  ],
  [
    "Summary, anywhere on the sheet",
    "No comment reading 'Goes away with its row'",
    "A comment on a row that was taken out went with it, instead of being left on someone else's row",
  ],
  [
    "Formulas, Name Manager",
    "Still lists Movements over Summary!$B$3:$D$5",
    "The workbook part is rewritten rather than copied, and the names came through it intact",
  ],
  [
    "Ledger rows 3 to 7",
    "Five entries, all banded and filtered like the table, with Total still above them",
    "A table with room for two grew to take five, and Excel still treats the new rows as part of it",
  ],
  [
    "Ledger!D1",
    "Reads 150, not 60",
    "SUM(Entries[Amount]) followed the table as it grew, which a formula written over a range could not",
  ],
  [
    "Click inside Ledger row 7",
    "The ribbon shows Table Design",
    "The last row we added really is inside the table, not just formatted like it",
  ],
  [
    "Ledger row 9",
    "Checked, and 150, where the template had them on row 6",
    "Growing the region pushed what was under it down by the three rows that arrived",
  ],
  [
    "Ledger!D9",
    "The formula bar shows SUM(D3:D7), not SUM(D3:D4)",
    "A total written over the region by range, rather than by the table's name, stretched to cover the rows that arrived",
  ],
  ["Ledger!B10", "Prepared by A. Person", "Plain text below the region moved down with everything else"],
];

async function buildTemplate() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Very Good Software";

  const report = workbook.addWorksheet("Report");
  report.views = [{ state: "frozen", ySplit: 8 }];
  report.getColumn(1).width = 34;
  report.getColumn(2).width = 10;
  report.getColumn(3).width = 16;

  report.mergeCells("A1:C1");
  report.getCell("A1").value = "ACME Consulting";
  report.getCell("A1").font = { bold: true, size: 14 };
  report.getCell("A1").fill = HEADER_FILL;

  report.getCell("A3").value = "Client";
  report.getCell("A4").value = "Issued";
  report.getCell("A5").value = "Due";

  // C4 is bold with no number format, so writing a date has to derive a format
  // and keep the bold. C5 already has one, so writing a date has to leave it be.
  report.getCell("C4").font = { bold: true };
  report.getCell("C5").numFmt = "dd/mm/yyyy";

  for (const [column, heading] of [
    ["A", "Item"],
    ["B", "Qty"],
    ["C", "Amount"],
  ]) {
    const cell = report.getCell(`${column}8`);
    cell.value = heading;
    cell.font = { bold: true };
    cell.fill = HEADER_FILL;
    cell.border = BORDERED;
  }

  // A data region formatted and waiting, but only two rows deep, so writing four
  // rows lands two on it and two past the end of it.
  for (const row of [9, 10]) {
    for (const column of ["A", "B", "C"]) {
      report.getCell(`${column}${row}`).border = BORDERED;
    }
    report.getCell(`C${row}`).numFmt = "#,##0.00";
  }

  report.getCell("A13").value = "Total";
  report.getCell("A13").font = { bold: true };
  // The cached result is deliberately wrong for the data about to be written, so
  // a file that did not recalculate shows 0 here.
  report.getCell("C13").value = { formula: "SUM(C9:C12)", result: 0 };
  report.getCell("C13").font = { bold: true };
  report.getCell("C13").numFmt = "#,##0.00";

  report.addConditionalFormatting({
    ref: "B9:B12",
    rules: [
      {
        type: "cellIs",
        operator: "greaterThan",
        formulae: [10],
        style: { font: { color: { argb: "FFCC0000" }, bold: true } },
      },
    ],
  });

  // A region the author named, filled with figures from a previous run. The fill
  // writes one row into it, so two rows have to go away and everything under them
  // has to come up, the total included.
  const summary = workbook.addWorksheet("Summary");
  summary.getCell("A1").value = "Movements by month";
  summary.getCell("A1").font = { bold: true };
  summary.getCell("A7").value = "Total";
  summary.getCell("A7").font = { bold: true };
  // One comment above the region, one on a row that goes away and one below it. A
  // comment is written across two parts, the cell it belongs to in one and the box
  // it appears in in another, so this is the only check that the pair agree.
  summary.getCell("A1").note = "Stays where it is";
  summary.getCell("A4").note = "Goes away with its row";
  summary.getCell("A7").note = "Comes up with the total";
  // Written over the region's rows, so shrinking the region has to rewrite it. The
  // cached result is wrong for the data about to arrive, as everywhere else here.
  summary.getCell("B7").value = { formula: "SUM(B3:B5)", result: 0 };
  summary.getCell("B7").font = { bold: true };
  // A rule below the region, naming a cell that moves. Both halves have to follow
  // the rows, the range it covers and the reference inside the rule, and the colour
  // is what says whether the second one did. Left where it was the rule reads B7,
  // which is empty once the total has come up, so the word comes out unhighlighted.
  summary.addConditionalFormatting({
    ref: "A7",
    rules: [{ type: "expression", formulae: ["$B7>0"], style: { font: { color: { argb: "FFCC0000" } } } }],
  });
  // The same question asked of a data validation, whose bounds are formulas of
  // their own. Its upper bound is the total, so a bound left behind bounds the cell
  // by an empty one and refuses everything above nothing.
  // Labelled beside itself, because a validation shows nothing until you type into
  // it, and a check you cannot find on the sheet is not a check.
  summary.getCell("D7").value = "<- type 50 here, then 500";
  summary.getCell("C7").border = BORDERED;
  summary.getCell("C7").dataValidation = {
    type: "whole",
    operator: "between",
    formulae: ["0", "$B7"],
    showErrorMessage: true,
    showInputMessage: true,
    promptTitle: "Data validation check",
    prompt: "50 should be accepted and 500 refused. Both refused means the upper bound stayed behind.",
    error: "Must be between 0 and the total",
  };
  for (const [row, month] of [
    [3, "January"],
    [4, "February"],
    [5, "March"],
  ]) {
    summary.getCell(`A${row}`).value = month;
    summary.getCell(`E${row}`).value = "keep me";
    for (const column of ["B", "C", "D"]) {
      const cell = summary.getCell(`${column}${row}`);
      cell.value = 999;
      cell.border = BORDERED;
      cell.numFmt = "#,##0";
    }
  }
  workbook.definedNames.add("Summary!B3:D5", "Movements");

  // A table with room for two rows and a total above it that refers to the table by
  // name rather than by range. The fill writes five rows, so the table has to grow
  // and the total has to follow it without being touched.
  const ledger = workbook.addWorksheet("Ledger");
  ledger.getCell("B1").value = "Total";
  ledger.getCell("B1").font = { bold: true };
  ledger.getCell("D1").value = { formula: "SUM(Entries[Amount])", result: 0 };
  ledger.getCell("D1").font = { bold: true };
  ledger.addTable({
    name: "Entries",
    ref: "B2",
    headerRow: true,
    columns: [{ name: "Entry" }, { name: "Qty" }, { name: "Amount" }],
    rows: [
      ["old one", 1, 999],
      ["old two", 2, 999],
    ],
  });

  // Anchored two rows below the total, which is itself below the region. The fill
  // takes two rows out of the region, so both the total and this have to come up by
  // two and the gap between them has to stay the same. exceljs cannot write a
  // chart, and a chart is anchored the same way, so an image stands in for one.
  const logo = workbook.addImage({ base64: PIXEL_PNG, extension: "png" });
  summary.addImage(logo, { tl: { col: 0, row: 8 }, ext: { width: 120, height: 60 } });

  // Below the table, so growing it has to push these down. The total is written
  // over the table's rows by range rather than by name, so it also has to stretch
  // to cover the rows that arrive.
  ledger.getCell("B6").value = "Checked";
  ledger.getCell("B6").font = { bold: true };
  ledger.getCell("D6").value = { formula: "SUM(D3:D4)", result: 0 };
  ledger.getCell("D6").font = { bold: true };
  ledger.getCell("B7").value = "Prepared by A. Person";

  const notes = workbook.addWorksheet("Notes");
  notes.getCell("A1").value = "This sheet is never opened by the fill, so every byte of it should survive.";

  return withCalculationChain(new Uint8Array(await workbook.xlsx.writeBuffer()));
}

// exceljs writes no calculation chain even for a workbook with formulas, so the
// path that drops one, along with its content type override and its relationship,
// would never be exercised. Excel does write one, so a real template has it, and
// leaving a dangling override or relationship behind is exactly the kind of thing
// that makes Excel offer to repair a file. Splicing one in gets that path covered.
function withCalculationChain(bytes) {
  const parts = unzipSync(bytes);

  parts["xl/calcChain.xml"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="${MAIN_NAMESPACE}"><c r="C13" i="1"/></calcChain>`,
  );
  parts["[Content_Types].xml"] = strToU8(
    strFromU8(parts["[Content_Types].xml"]).replace(
      "</Types>",
      `<Override PartName="/xl/calcChain.xml" ContentType="${SPREADSHEET_TYPES}.calcChain+xml"/></Types>`,
    ),
  );
  parts["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(parts["xl/_rels/workbook.xml.rels"]).replace(
      "</Relationships>",
      `<Relationship Id="rIdCalcChain" Type="${CALC_CHAIN_RELATIONSHIP}" Target="calcChain.xml"/></Relationships>`,
    ),
  );

  return zipSync(parts);
}

async function fill(source) {
  const editor = (await Workbook.open(source)).edit();
  const report = editor.worksheet("Report");

  report.set("C3", "Northwind Traders");
  report.set("C4", date("2026-03-01"));
  report.set("C5", date("2026-03-15"));

  // Four rows into a region two rows deep, so rows 11 and 12 are new and take
  // their formatting from row 9.
  report.writeRows(
    9,
    [
      ["Discovery workshop", 4, 400],
      ["Implementation", 12, 600],
      ["Review", 2, 240],
      ["Handover", 1, 160],
    ],
    { inheritFrom: 9 },
  );

  // One row into a region three rows deep, so the other two are cleared and the
  // labels either side of it are left alone.
  editor.worksheet("Summary").writeRegion("Movements", [[120, 45, 165]]);

  // Five rows into a table with room for two, so it has to grow by three.
  editor.worksheet("Ledger").writeRegion("Entries", [
    ["Discovery", 4, 10],
    ["Build", 12, 20],
    ["Review", 2, 30],
    ["Handover", 1, 40],
    ["Support", 3, 50],
  ]);

  const checks = editor.addWorksheet("Checks");
  checks.appendRows(EXPECTATIONS);

  return editor.save();
}

// The pivot template holds five rows over three regions. Eight rows arrive, one of
// them for a region the cache has never seen, so West appearing in the summary is
// Excel having rebuilt the cache rather than redrawn the old one.
const PIVOT_ROWS = [
  ["North", "January", 100],
  ["North", "February", 100],
  ["South", "January", 100],
  ["South", "February", 100],
  ["East", "January", 100],
  ["East", "February", 100],
  ["West", "January", 100],
  ["West", "February", 100],
];

const PIVOT_EXPECTATIONS = [
  ["Where", "What you should see", "Why it matters"],
  ["The file itself", "Opens with no repair prompt", "A rewritten pivot cache is accepted"],
  [
    "Data rows 2 to 9",
    "Eight rows, each with its own region and month, replacing the five that were there",
    "The region took the rows it was given, and every row kept its own values rather than repeating the last",
  ],
  [
    "The Pivot sheet",
    "Four regions listed, East, North, South and West",
    "West is in none of the cached rows, so seeing it means Excel rebuilt the cache from the range rather than redrawing what it had",
  ],
  [
    "The Pivot sheet's grand total",
    "Reads 800, not 1000",
    "1000 is the total of the rows the file was saved with, so that figure means the cache was not rebuilt",
  ],
  [
    "The pivot's source range",
    "Click a cell in the pivot, open the PivotTable Analyze tab that appears, click Change Data Source. It should read Data!$A$1:$C$9. Press Cancel, not OK",
    "The only check that reads the range we wrote rather than inferring it from what the pivot drew. Older Excel calls that tab Analyze or Options, and if you cannot find it at all, skip this row, since the West check covers the same ground",
  ],
];

async function fillPivot(source) {
  const editor = (await Workbook.open(source)).edit();

  editor.writeRegion("Movements", PIVOT_ROWS);
  editor.addWorksheet("Checks").appendRows(PIVOT_EXPECTATIONS);

  return editor.save();
}

// The chart template holds five rows over `Alpha` to `Epsilon`, all at 100, and
// every chart in it carries a cached copy of exactly those. Eight rows arrive under
// names none of them has, so a bar labelled `Row 1` is Excel having read the range
// we rewrote rather than redrawing what the chart already held.
const CHART_ROWS = [1, 2, 3, 4, 5, 6, 7, 8].map((n) => [`Row ${n}`, n, 50]);

const CHART_EXPECTATIONS = [
  ["Where", "What you should see", "Why it matters"],
  ["The file itself", "Opens with no repair prompt", "A rewritten chart part is accepted"],
  [
    "Data rows 2 to 9",
    "Eight rows, Row 1 to Row 8, replacing the five that were there",
    "The region took the rows it was given",
  ],
  [
    "Data!A11",
    "Reads Total, where the template had it at A8",
    "Everything under the region went down by the three rows that arrived",
  ],
  [
    "Data!C11",
    "Reads 400, and the formula bar shows SUM(C2:C9)",
    "The total was written over the region's rows and stretched to cover the ones that arrived",
  ],
  [
    "The chart on the Data sheet",
    "Eight bars, labelled Row 1 to Row 8. Not five labelled Alpha to Epsilon",
    "Every chart holds a cached copy of the rows it read, and Alpha to Epsilon is all that copy has, so new labels mean Excel read the range we rewrote instead of redrawing what it held",
  ],
  [
    "The same chart, where it sits",
    "Still one empty row below Total, now starting on row 13",
    "A chart anchored below the region came down with the rows. An anchor that did not move leaves it starting on row 10, which is above Total rather than below it",
  ],
  [
    "The chart on the Dashboard sheet",
    "The same eight bars",
    "Dashboard has no rows of its own that moved, so this one proves a series follows what it reads rather than where it is drawn",
  ],
  [
    "The Chart1 tab",
    "The same eight bars",
    "A chart on its own tab hangs off no worksheet, so this one proves every chart in the file was found rather than only the ones a worksheet points at",
  ],
  [
    "Any of the three charts",
    "Click the chart, open the Chart Design tab that appears, click Select Data. The Chart data range box at the top should read =Data!$A$1:$C$9, where the template had =Data!$A$1:$C$6. Press Cancel",
    "The only check that reads the reference we wrote rather than inferring it from what the chart drew. Older Excel puts Select Data under the Design or Chart Design tab, and if you cannot find it, skip this row, since the Row 1 labels cover the same ground",
  ],
];

async function fillChart(source) {
  const editor = (await Workbook.open(source)).edit();

  editor.writeRegion("Movements", CHART_ROWS);
  editor.addWorksheet("Checks").appendRows(CHART_EXPECTATIONS);

  return editor.save();
}

async function collect(stream) {
  const chunks = [];
  const reader = stream.getReader();

  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    chunks.push(value);
  }

  return Buffer.concat(chunks);
}

mkdirSync(OUTPUT_DIR, { recursive: true });

const provided = existsSync(TEMPLATE);
if (!provided) {
  writeFileSync(TEMPLATE, await buildTemplate());
}

const source = readFileSync(TEMPLATE);
writeFileSync(FILLED, await collect(await fill(source)));
writeFileSync(PIVOT_FILLED, await collect(await fillPivot(readFileSync(PIVOT_TEMPLATE))));
writeFileSync(CHART_FILLED, await collect(await fillChart(readFileSync(CHART_TEMPLATE))));

console.log(`template: ${TEMPLATE}${provided ? " (yours)" : " (generated)"}`);
console.log(`filled:   ${FILLED}`);
console.log(`pivot:    ${PIVOT_FILLED} (from ${PIVOT_TEMPLATE})`);
console.log(`chart:    ${CHART_FILLED} (from ${CHART_TEMPLATE})`);
console.log();
console.log("Open all three filled files in Excel and work down each one's Checks sheet.");
