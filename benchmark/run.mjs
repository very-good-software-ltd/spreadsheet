// Runs the worker once per (library, file) in its own process, so no library's
// memory bleeds into another's measurement, and prints a table. With no file
// arguments it runs every .xlsx in benchmark/files, skipping names that start
// with an underscore. Those files are gitignored, so drop any files in there to
// play with, and prefix a name with _ to leave it out of a run. Pass --cap=512
// to run the workers under a heap cap to see which libraries run out of memory.
import { spawnSync } from "node:child_process";
import { readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

const FILES_DIR = "benchmark/files";
const LIBRARIES = ["very-good-spreadsheet", "exceljs", "xlsx"];
const worker = fileURLToPath(new URL("./read-file.mjs", import.meta.url));

const args = process.argv.slice(2);
const capArg = args.find((arg) => arg.startsWith("--cap="));
const cap = capArg ? Number(capArg.slice("--cap=".length)) : undefined;
const explicit = args.filter((arg) => !arg.startsWith("--"));
const targets = explicit.length > 0 ? explicit : discoverFiles();

function discoverFiles() {
  let names;
  try {
    names = readdirSync(FILES_DIR);
  } catch {
    return [];
  }
  return names.filter((name) => name.endsWith(".xlsx") && !name.startsWith("_")).map((name) => `${FILES_DIR}/${name}`);
}

if (targets.length === 0) {
  console.log(`No .xlsx files in ${FILES_DIR}. Add some there, or pass file paths as arguments.`);
  process.exit(0);
}

function run(library, file) {
  const nodeArgs = [...(cap ? [`--max-old-space-size=${cap}`] : []), worker, library, file];
  const result = spawnSync(process.execPath, nodeArgs, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  if (result.status === 0) {
    return { ok: true, ...JSON.parse(result.stdout.trim().split("\n").at(-1)) };
  }
  const ranOutOfMemory = /heap out of memory/i.test(result.stderr) || result.signal === "SIGABRT";
  return { ok: false, status: ranOutOfMemory ? "out of memory" : `failed (${result.signal ?? result.status})` };
}

function pad(value, width) {
  return String(value).padEnd(width);
}

function padStart(value, width) {
  return String(value).padStart(width);
}

for (const file of targets) {
  console.log(`\n${file}${cap ? `  (heap cap ${cap}MB)` : ""}`);
  console.log(`  ${pad("library", 24)} ${padStart("cells", 10)} ${padStart("time", 9)} ${padStart("peak RSS", 10)}`);
  for (const library of LIBRARIES) {
    const outcome = run(library, file);
    if (outcome.ok) {
      console.log(
        `  ${pad(library, 24)} ${padStart(outcome.cells, 10)} ${padStart(`${(outcome.ms / 1000).toFixed(1)}s`, 9)} ${padStart(`${outcome.rssMb.toFixed(0)}MB`, 10)}`,
      );
    } else {
      console.log(`  ${pad(library, 24)} ${padStart(outcome.status, 32)}`);
    }
  }
}
