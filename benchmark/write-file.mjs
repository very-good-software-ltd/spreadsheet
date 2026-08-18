// Writes a file of N rows with one library in one mode, then prints a JSON line
// with the time, the size of the file and the OS-tracked peak resident memory.
// Run as its own process, one per (library, mode, count), so each measurement
// starts clean.
//
// Every mode writes to a real file on disk, so nothing accumulates on the
// consumer's side and what is measured is the library rather than the sink.
//
// Modes:
//   stream - hand the rows over as a source and let the writer pull them
//   load   - build every row in memory first, the load-everything path
//   region      - fill a template's named region from an array, which moves the
//                 rows below it. Only we have this, so it has no counterpart in the
//                 other libraries.
//   region-lazy - the same from a source whose length is not known up front, which
//                 is the one write path that has to hold every row it is given.
import { createWriteStream, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { performance } from "node:perf_hooks";
import { Writable } from "node:stream";

async function importDefault(specifier) {
  const module = await import(specifier);
  return module.default ?? module;
}

// One shape for every library, so the comparison is of writers and not of row
// shapes. Four columns covering the types a real export mixes.
function rowAt(index) {
  return [index, `Order ${index}`, index * 1.5, index % 7 === 0];
}

function* rows(count) {
  for (let index = 1; index <= count; index += 1) {
    yield rowAt(index);
  }
}

// A one-row region under a heading, which is the shape of a real template: the
// author draws one row to show what a row looks like and names it.
async function regionTemplate() {
  const ExcelJS = await importDefault("exceljs");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Data");
  sheet.addRow(["Id", "Description", "Amount", "Flagged"]);
  sheet.addRow(rowAt(0));
  workbook.definedNames.add("Data!A2:D2", "Data");

  return new Uint8Array(await workbook.xlsx.writeBuffer());
}

const runners = {
  "@very-good-software/spreadsheet": {
    async stream(count, path) {
      const { Workbook } = await importDefault("../dist/index.js");
      const editor = (await Workbook.create()).edit();
      editor.worksheet(0).appendRows(rows(count));
      await editor.save().pipeTo(Writable.toWeb(createWriteStream(path)));
    },
    async load(count, path) {
      const { Workbook } = await importDefault("../dist/index.js");
      const editor = (await Workbook.create()).edit();
      // Materialising the rows first is the thing the streaming mode avoids, so
      // the pair shows what handing over a source rather than an array buys.
      editor.worksheet(0).appendRows([...rows(count)]);
      await editor.save().pipeTo(Writable.toWeb(createWriteStream(path)));
    },
    async region(count, path) {
      const { Workbook } = await importDefault("../dist/index.js");
      const editor = (await Workbook.open(await regionTemplate())).edit();
      // An array says how many rows it has, so nothing has to be held to find out.
      editor.worksheet(0).writeRegion("Data", [...rows(count)]);
      await editor.save().pipeTo(Writable.toWeb(createWriteStream(path)));
    },
    async "region-lazy"(count, path) {
      const { Workbook } = await importDefault("../dist/index.js");
      const editor = (await Workbook.open(await regionTemplate())).edit();
      // A source that only says how many rows it had once it is finished, so every
      // row is held while the count is worked out.
      editor.worksheet(0).writeRegion("Data", rows(count));
      await editor.save().pipeTo(Writable.toWeb(createWriteStream(path)));
    },
  },
  exceljs: {
    async stream(count, path) {
      const ExcelJS = await importDefault("exceljs");
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path });
      const sheet = workbook.addWorksheet("Sheet1");
      for (const row of rows(count)) {
        sheet.addRow(row).commit();
      }
      sheet.commit();
      await workbook.commit();
    },
    async load(count, path) {
      const ExcelJS = await importDefault("exceljs");
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Sheet1");
      for (const row of rows(count)) {
        sheet.addRow(row);
      }
      await workbook.xlsx.writeFile(path);
    },
  },
  xlsx: {
    async load(count, path) {
      const XLSX = await importDefault("xlsx");
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([...rows(count)]), "Sheet1");
      // SheetJS stores entries uncompressed unless told otherwise, which would put
      // it in the size column against everyone else's deflated output.
      XLSX.writeFile(workbook, path, { compression: true });
    },
  },
};

const [, , library, mode, countText] = process.argv;
const count = Number(countText);
const runner = runners[library]?.[mode];
if (runner === undefined) {
  console.error(`No runner for ${library}/${mode}`);
  process.exit(2);
}

const path = join(tmpdir(), `very-good-spreadsheet-benchmark-${process.pid}.xlsx`);

try {
  const start = performance.now();
  await runner(count, path);
  const ms = performance.now() - start;

  // Node reports maxRSS in kilobytes, and it is the lifetime peak, so a mode that
  // built everything first is still credited with its high-water mark.
  const rssMb = (process.resourceUsage().maxRSS * 1024) / (1024 * 1024);
  const bytes = statSync(path).size;

  console.log(JSON.stringify({ library, mode, rows: count, bytes, ms, rssMb }));
} finally {
  rmSync(path, { force: true });
}
