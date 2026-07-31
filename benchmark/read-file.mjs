// Reads every cell of one file with one library, then prints a JSON line with
// the time and the OS-tracked peak resident memory. Run as its own process, one
// per (library, file), so each measurement starts from a clean Node baseline.
import { readFileSync } from "node:fs";
import { performance } from "node:perf_hooks";

async function load(specifier) {
  const module = await import(specifier);
  return module.default ?? module;
}

async function readVeryGoodSpreadsheet(path) {
  const { Workbook } = await load("../dist/index.js");
  const workbook = await Workbook.open(readFileSync(path));
  let cells = 0;
  for (const name of workbook.worksheetNames) {
    for await (const row of workbook.worksheet(name).rows()) {
      cells += row.cells.length;
    }
  }
  return cells;
}

async function readExcelJs(path) {
  const ExcelJS = await load("exceljs");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path);
  let cells = 0;
  workbook.eachSheet((sheet) => {
    sheet.eachRow((row) => {
      row.eachCell(() => {
        cells += 1;
      });
    });
  });
  return cells;
}

async function readSheetJs(path) {
  const XLSX = await load("xlsx");
  const workbook = XLSX.readFile(path, { cellFormula: false, cellHTML: false });
  let cells = 0;
  for (const name of workbook.SheetNames) {
    for (const key of Object.keys(workbook.Sheets[name])) {
      if (!key.startsWith("!")) {
        cells += 1;
      }
    }
  }
  return cells;
}

const readers = {
  "very-good-spreadsheet": readVeryGoodSpreadsheet,
  exceljs: readExcelJs,
  xlsx: readSheetJs,
};

const [, , library, file] = process.argv;
const reader = readers[library];
if (reader === undefined) {
  console.error(`Unknown library: ${library}`);
  process.exit(2);
}

const start = performance.now();
const cells = await reader(file);
const ms = performance.now() - start;

// Node reports maxRSS in kilobytes.
const rssBytes = process.resourceUsage().maxRSS * 1024;

console.log(JSON.stringify({ library, file, cells, ms, rssMb: rssBytes / (1024 * 1024) }));
