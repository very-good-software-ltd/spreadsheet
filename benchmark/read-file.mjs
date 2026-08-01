// Reads every cell of one file with one library in one mode, then prints a JSON
// line with the time and the OS-tracked peak resident memory. Run as its own
// process, one per (library, mode, file), so each measurement starts clean.
//
// Modes:
//   stream - read row by row and discard, the bounded-memory path
//   load   - materialize every row and hold it, the load-everything path
import { readFileSync } from "node:fs";
import { performance } from "node:perf_hooks";

async function importDefault(specifier) {
  const module = await import(specifier);
  return module.default ?? module;
}

const runners = {
  "very-good-spreadsheet": {
    async stream(file) {
      const { Workbook } = await importDefault("../dist/index.js");
      const workbook = await Workbook.open(readFileSync(file));
      let cells = 0;
      for (const name of workbook.worksheetNames) {
        for await (const row of workbook.worksheet(name).rows()) {
          cells += row.cells.length;
        }
      }
      return cells;
    },
    async load(file) {
      const { Workbook } = await importDefault("../dist/index.js");
      const workbook = await Workbook.open(readFileSync(file));
      const sheets = [];
      for (const name of workbook.worksheetNames) {
        sheets.push(await Array.fromAsync(workbook.worksheet(name).rows()));
      }
      return sheets.reduce((sum, rows) => sum + rows.reduce((n, row) => n + row.cells.length, 0), 0);
    },
  },
  exceljs: {
    async stream(file) {
      const ExcelJS = await importDefault("exceljs");
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, {});
      let cells = 0;
      for await (const worksheet of reader) {
        for await (const row of worksheet) {
          row.eachCell(() => {
            cells += 1;
          });
        }
      }
      return cells;
    },
    async load(file) {
      const ExcelJS = await importDefault("exceljs");
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(file);
      let cells = 0;
      workbook.eachSheet((sheet) => {
        sheet.eachRow((row) => {
          row.eachCell(() => {
            cells += 1;
          });
        });
      });
      return cells;
    },
  },
  xlsx: {
    async load(file) {
      const XLSX = await importDefault("xlsx");
      const workbook = XLSX.readFile(file, { cellFormula: false, cellHTML: false });
      let cells = 0;
      for (const name of workbook.SheetNames) {
        for (const key of Object.keys(workbook.Sheets[name])) {
          if (!key.startsWith("!")) {
            cells += 1;
          }
        }
      }
      return cells;
    },
  },
};

const [, , library, mode, file] = process.argv;
const runner = runners[library]?.[mode];
if (runner === undefined) {
  console.error(`No runner for ${library}/${mode}`);
  process.exit(2);
}

const start = performance.now();
const cells = await runner(file);
const ms = performance.now() - start;

// Node reports maxRSS in kilobytes; it is the lifetime peak, so the retained
// load mode's high-water mark is captured even after the data is dropped.
const rssMb = (process.resourceUsage().maxRSS * 1024) / (1024 * 1024);

console.log(JSON.stringify({ library, mode, file, cells, ms, rssMb }));
