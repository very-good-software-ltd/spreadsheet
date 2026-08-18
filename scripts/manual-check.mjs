// Produces the pair of files the manual checks in MANUAL-CHECKS.md need: a
// template, and that template filled in by this library. Open the filled one in
// Excel and work down its "Checks" sheet, which lists what to look at and what
// you should see.
//
// Drop your own file at manual-check/template.xlsx and it is used instead of the
// generated one. Do that for the checks nothing here can cover, charts and pivot
// tables, which exceljs cannot write.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { date, Workbook } from "../dist/index.js";

const OUTPUT_DIR = "manual-check";
const TEMPLATE = join(OUTPUT_DIR, "template.xlsx");
const FILLED = join(OUTPUT_DIR, "filled.xlsx");

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
    "The image on Summary",
    "Still two rows below Total, at about row 7",
    "A picture anchored below the region came up with the rows. Four rows of gap means it did not move, and a chart is anchored the same way",
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
  // Written over the region's rows, so shrinking the region has to rewrite it. The
  // cached result is wrong for the data about to arrive, as everywhere else here.
  summary.getCell("B7").value = { formula: "SUM(B3:B5)", result: 0 };
  summary.getCell("B7").font = { bold: true };
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

console.log(`template: ${TEMPLATE}${provided ? " (yours)" : " (generated)"}`);
console.log(`filled:   ${FILLED}`);
console.log();
console.log("Open the filled file in Excel and work down its Checks sheet.");
if (!provided) {
  console.log();
  console.log("Charts and pivot tables are not covered: exceljs cannot write them.");
  console.log(`For those, put a real template at ${TEMPLATE} and run this again.`);
}
