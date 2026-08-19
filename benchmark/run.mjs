// Runs the worker once per (library, file) in its own process, so no library's
// memory bleeds into another's measurement, and prints a table.
//
//   npm run benchmark          reading and writing
//   npm run benchmark:read     reading only
//   npm run benchmark:write    writing only
//
// Reading covers every .xlsx and .ods in benchmark/files, skipping names that
// start with an underscore. Those files are gitignored, so drop any files in there
// to play with, and prefix a name with _ to leave it out of a run. Pass file paths
// as arguments to read those instead.
//
// Writing is a million rows by default. Pass --write=100000 for a quicker loop, or
// a list like --write=100000,1000000 to see whether peak memory moves with the row
// count.
//
// Pass --cap=150 to run the workers under a heap cap. That is the measure worth
// quoting for memory, since peak RSS also counts memory the runtime has not given
// back to the OS, and it shows which libraries run out.
import { spawnSync } from "node:child_process";
import { readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

const FILES_DIR = "benchmark/files";
// Each library in every mode it supports. stream is the bounded-memory path;
// load materializes everything. SheetJS has no streaming read, and exceljs reads
// xlsx only, so it is skipped for .ods files.
const RUNS = [
  ["@very-good-software/spreadsheet", "stream"],
  ["@very-good-software/spreadsheet", "load"],
  ["exceljs", "stream"],
  ["exceljs", "load"],
  ["xlsx", "load"],
];

// A grid: append rows or fill a region, from a source or from an array. SheetJS
// builds the whole sheet before writing, and exceljs has a streaming writer as well
// as its ordinary one, the same split as on the read side. The two region modes are
// ours alone, since filling a named region moves the rows below it.
const WRITE_RUNS = [
  ["@very-good-software/spreadsheet", "stream"],
  ["@very-good-software/spreadsheet", "load"],
  ["@very-good-software/spreadsheet", "region-stream"],
  ["@very-good-software/spreadsheet", "region-load"],
  ["exceljs", "stream"],
  ["exceljs", "load"],
  ["xlsx", "load"],
];

function runsFor(file) {
  return file.endsWith(".ods") ? RUNS.filter(([library]) => library !== "exceljs") : RUNS;
}
const reader = fileURLToPath(new URL("./read-file.mjs", import.meta.url));
const writer = fileURLToPath(new URL("./write-file.mjs", import.meta.url));

const DEFAULT_ROW_COUNT = 1_000_000;

const args = process.argv.slice(2);
const capArg = args.find((arg) => arg.startsWith("--cap="));
const cap = capArg ? Number(capArg.slice("--cap=".length)) : undefined;

const writeArg = args.find((arg) => arg === "--write" || arg.startsWith("--write="));
const readArg = args.some((arg) => arg === "--read");
const explicit = args.filter((arg) => !arg.startsWith("--"));

// Neither asked for means both, so running it bare shows everything rather than
// half of it. Either one asked for means only that one.
const reading = readArg || explicit.length > 0 || writeArg === undefined;
const writing = writeArg !== undefined || !readArg;

const rowCounts = !writing
  ? []
  : writeArg === undefined || writeArg === "--write"
    ? [DEFAULT_ROW_COUNT]
    : writeArg
        .slice("--write=".length)
        .split(",")
        .map((count) => Number(count));
const targets = !reading ? [] : explicit.length > 0 ? explicit : discoverFiles();

function discoverFiles() {
  let names;
  try {
    names = readdirSync(FILES_DIR);
  } catch {
    return [];
  }
  return names
    .filter((name) => (name.endsWith(".xlsx") || name.endsWith(".ods")) && !name.startsWith("_"))
    .map((name) => `${FILES_DIR}/${name}`);
}

if (reading && targets.length === 0 && rowCounts.length === 0) {
  console.log(`No .xlsx or .ods files in ${FILES_DIR}. Add some there, or pass file paths as arguments.`);
  console.log("To benchmark writing instead, pass --write=1000000.");
  process.exit(0);
}

function run(worker, library, mode, target) {
  const nodeArgs = [...(cap ? [`--max-old-space-size=${cap}`] : []), worker, library, mode, String(target)];
  const result = spawnSync(process.execPath, nodeArgs, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  if (result.status === 0) {
    return { ok: true, ...JSON.parse(result.stdout.trim().split("\n").at(-1)) };
  }
  const ranOutOfMemory = /heap out of memory/i.test(result.stderr) || result.signal === "SIGABRT";
  if (ranOutOfMemory) {
    return { ok: false, status: "out of memory" };
  }
  return { ok: false, status: "failed", reason: errorReason(result.stderr) };
}

function errorReason(stderr) {
  const match = stderr.match(/^\s*(?:[A-Z]\w*Error|Error): .*/m);
  return (match ? match[0] : (stderr.trim().split("\n").at(-1) ?? "")).trim();
}

function pad(value, width) {
  return String(value).padEnd(width);
}

function padStart(value, width) {
  return String(value).padStart(width);
}

function report(outcome, label, measure) {
  if (outcome.ok) {
    console.log(
      `${label} ${padStart(measure(outcome), 10)} ${padStart(`${(outcome.ms / 1000).toFixed(1)}s`, 9)} ${padStart(`${outcome.rssMb.toFixed(0)}MB`, 10)}`,
    );
    return;
  }
  console.log(`${label} ${padStart(outcome.status, 31)}`);
  if (outcome.reason) {
    console.log(`  ${outcome.reason}`);
  }
}

for (const file of targets) {
  console.log(`\n${file}${cap ? `  (heap cap ${cap}MB)` : ""}`);
  console.log(
    `  ${pad("library", 35)} ${pad("mode", 14)} ${padStart("cells", 10)} ${padStart("time", 9)} ${padStart("peak RSS", 10)}`,
  );
  for (const [library, mode] of runsFor(file)) {
    report(run(reader, library, mode, file), `  ${pad(library, 35)} ${pad(mode, 14)}`, (outcome) => outcome.cells);
  }
}

for (const count of rowCounts) {
  console.log(`\nwriting ${count.toLocaleString("en-GB")} rows${cap ? `  (heap cap ${cap}MB)` : ""}`);
  console.log(
    `  ${pad("library", 35)} ${pad("mode", 14)} ${padStart("file", 10)} ${padStart("time", 9)} ${padStart("peak RSS", 10)}`,
  );
  for (const [library, mode] of WRITE_RUNS) {
    report(
      run(writer, library, mode, count),
      `  ${pad(library, 35)} ${pad(mode, 14)}`,
      (outcome) => `${(outcome.bytes / 1e6).toFixed(1)}MB`,
    );
  }
}
