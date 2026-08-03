import { toByteRange } from "./io/byte-range";
import type { BinarySource } from "./io/source";
import type { Row } from "./row";
import { Worksheet } from "./worksheet";
import { readXlsx } from "./xlsx/read-xlsx";

// A whole-bytes input, held in memory, or a seekable Blob, a File in the browser
// or an fs.openAsBlob handle in Node, read in ranges so the file is never fully
// held.
export type WorkbookSource = BinarySource | Blob;

export interface WorksheetInfo {
  readonly name: string;
  readonly hidden: boolean;
}

// What a format reader hands the workbook: the sheets in document order, and a
// way to stream one sheet's rows by position. Each format we read produces this.
export interface WorkbookData {
  readonly worksheets: readonly WorksheetInfo[];
  openRows(index: number): AsyncIterable<Row>;
}

export class Workbook {
  static async open(source: WorkbookSource): Promise<Workbook> {
    return new Workbook(await readXlsx(await toByteRange(source)));
  }

  constructor(private readonly data: WorkbookData) {}

  get worksheets(): readonly WorksheetInfo[] {
    return this.data.worksheets;
  }

  get worksheetNames(): readonly string[] {
    return this.data.worksheets.map((sheet) => sheet.name);
  }

  worksheet(nameOrIndex: string | number): Worksheet {
    const index =
      typeof nameOrIndex === "number"
        ? nameOrIndex
        : this.data.worksheets.findIndex((sheet) => sheet.name === nameOrIndex);

    if (index < 0 || index >= this.data.worksheets.length) {
      throw new Error(`Worksheet not found: ${nameOrIndex}`);
    }

    return new Worksheet(() => this.data.openRows(index));
  }

  firstWorksheet(): Worksheet {
    if (this.data.worksheets.length === 0) {
      throw new Error("Workbook has no worksheets");
    }

    return new Worksheet(() => this.data.openRows(0));
  }
}
