export const REPO_URL = "https://github.com/christophgockel/very-good-spreadsheet";
export const NPM_URL = "https://www.npmjs.com/package/@very-good-software/spreadsheet";
export const NPM_INSTALL = "npm install @very-good-software/spreadsheet";

export interface Feature {
  readonly title: string;
  readonly body: string;
}

export const FEATURES: readonly Feature[] = [
  {
    title: "Streaming, low memory",
    body: "It streams the sheet instead of expanding it in memory, so memory stays flat as the sheet grows. Files that need gigabytes in other libraries read here in a couple hundred megabytes.",
  },
  {
    title: "Node and the browser",
    body: "One codebase on web standards. Other libraries stream in Node, but this streams a real file in the browser too, with no server.",
  },
  {
    title: "Typed cells",
    body: "A number comes back as a number and a date as a Date, so you never have to parse a cell yourself.",
  },
  {
    title: "Fast",
    body: "On xlsx it's the fastest reader in our benchmark, ahead of the libraries that load the whole file.",
  },
  {
    title: "Tiny",
    body: "One runtime dependency, an ES module, with TypeScript types included.",
  },
];

export interface BenchmarkRow {
  readonly library: string;
  readonly mode: "streaming" | "loading";
  readonly timeSeconds: number;
  readonly peakMemoryMb: number;
  readonly us?: boolean;
}

// Reading every cell of a 28 MB file whose single sheet is 170 MB uncompressed
// (4.44M cells), each library in its own Node process. Reproduce with `npm run benchmark`.
export const BENCHMARK: readonly BenchmarkRow[] = [
  { library: "@very-good-software/spreadsheet", mode: "streaming", timeSeconds: 3.4, peakMemoryMb: 145, us: true },
  { library: "exceljs", mode: "streaming", timeSeconds: 4.3, peakMemoryMb: 219 },
  { library: "SheetJS (xlsx)", mode: "loading", timeSeconds: 8.3, peakMemoryMb: 1503 },
  { library: "exceljs", mode: "loading", timeSeconds: 10.2, peakMemoryMb: 2756 },
];

export interface OdsBenchmarkRow {
  readonly file: string;
  readonly cells: string;
  readonly ours: string;
  readonly sheetjs: string;
}

// Reading every cell of an .ods file, our streaming read against SheetJS load.
// exceljs cannot read .ods at all. Reproduce with `npm run benchmark`.
export const ODS_BENCHMARK: readonly OdsBenchmarkRow[] = [
  { file: "5 MB", cells: "1.2M cells", ours: "7.3s, 131 MB", sheetjs: "5.9s, 1.3 GB" },
  { file: "14 MB", cells: "2.6M cells", ours: "31s, 201 MB", sheetjs: "13s, 2.5 GB" },
  { file: "48 MB", cells: "10.7M cells", ours: "61s, 249 MB", sheetjs: "crashes" },
];
