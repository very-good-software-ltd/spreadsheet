export const REPO_URL = "https://github.com/christophgockel/very-good-spreadsheet";
export const NPM_INSTALL = "npm install very-good-spreadsheet";

export interface Feature {
  readonly title: string;
  readonly body: string;
}

export const FEATURES: readonly Feature[] = [
  {
    title: "Streaming, low memory",
    body: "It streams the sheet instead of expanding it in memory, so memory stays flat as the sheet grows. A file that costs other libraries gigabytes reads here in a couple hundred megabytes.",
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
    body: "The fastest and leanest reader in our benchmark, ahead of the libraries that load the whole file.",
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
  { library: "very-good-spreadsheet", mode: "streaming", timeSeconds: 3.4, peakMemoryMb: 145, us: true },
  { library: "exceljs", mode: "streaming", timeSeconds: 4.3, peakMemoryMb: 219 },
  { library: "SheetJS (xlsx)", mode: "loading", timeSeconds: 8.3, peakMemoryMb: 1503 },
  { library: "exceljs", mode: "loading", timeSeconds: 10.2, peakMemoryMb: 2756 },
];
