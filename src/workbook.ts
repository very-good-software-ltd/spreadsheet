import type { Editor } from "./editor";
import { toByteRange } from "./io/byte-range";
import type { BinarySource } from "./io/source";
import { readSpreadsheet } from "./read-spreadsheet";
import type { Row } from "./row";
import { Worksheet } from "./worksheet";
import { blankXlsxArchive } from "./xlsx/blank-workbook";
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

  /** An editor over the same file. Absent for a format that only reads. */
  readonly edit?: () => Editor;
}

export class Workbook {
  static async open(source: WorkbookSource): Promise<Workbook> {
    return new Workbook(await readSpreadsheet(await toByteRange(source)));
  }

  /**
   * A workbook with one empty worksheet, as the starting point for building a
   * file. It behaves exactly like an opened one, so creating a file from scratch
   * and filling a template are the same code from here on.
   */
  static async create(): Promise<Workbook> {
    return new Workbook(await readXlsx(blankXlsxArchive()));
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

  /**
   * An editor that writes a new file from this one. Reading and writing are
   * separate: the editor never reflects edits back into this workbook's rows.
   *
   * Throws for a format we do not write.
   */
  edit(): Editor {
    if (this.data.edit === undefined) {
      throw new Error("Writing this format is not supported");
    }
    return this.data.edit();
  }

  firstWorksheet(): Worksheet {
    if (this.data.worksheets.length === 0) {
      throw new Error("Workbook has no worksheets");
    }

    return new Worksheet(() => this.data.openRows(0));
  }
}
