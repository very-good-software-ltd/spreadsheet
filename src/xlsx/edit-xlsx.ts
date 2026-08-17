import type { CellInput } from "../cell-input";
import { columnIndexOf, rowNumberOf } from "../cell-reference";
import type { Editor, RowSource, WorksheetEditor, WriteRowsOptions } from "../editor";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlReader } from "../xml/xml-reader";
import { createZipWriter } from "../zip/create-zip-writer";
import type { ZipArchive } from "../zip/zip-archive";
import { type CellEdit, mergeRowEdits, type RowBlock } from "./merge-row-edits";
import type { Styles } from "./read-styles";
import type { WorksheetRef } from "./read-workbook";
import { writeSheetPart } from "./write-sheet";
import { DateStyleTable, writeStylesPart } from "./write-styles";

const STYLES_PART = "xl/styles.xml";

const CELL_REFERENCE = /^[A-Z]{1,3}[1-9][0-9]*$/;

// The last row and column a worksheet can hold.
const MAX_ROW = 1_048_576;
const MAX_COLUMN_INDEX = 16_383;

interface SheetEdits {
  readonly cells: CellEdit[];
  readonly blocks: RowBlock[];
  readonly appended: RowSource[];
}

export class XlsxEditor implements Editor {
  private readonly edits = new Map<number, SheetEdits>();
  private readonly xml: XmlReader = createXmlReader();
  private calls = 0;
  private saved = false;

  constructor(
    private readonly archive: ZipArchive,
    private readonly worksheets: readonly WorksheetRef[],
    private readonly styles: Styles,
    private readonly date1904: boolean,
  ) {}

  worksheet(nameOrIndex: string | number): WorksheetEditor {
    const index =
      typeof nameOrIndex === "number" ? nameOrIndex : this.worksheets.findIndex((sheet) => sheet.name === nameOrIndex);

    if (index < 0 || index >= this.worksheets.length) {
      throw new Error(`Worksheet not found: ${nameOrIndex}`);
    }

    return new XlsxWorksheetEditor(this.editsFor(index), () => {
      this.calls += 1;
      return this.calls;
    });
  }

  save(): ReadableStream<Uint8Array> {
    if (this.saved) {
      throw new Error(
        "This workbook has already been saved. Row sources are read once, so saving again would drop rows",
      );
    }
    this.saved = true;

    const dateStyles = new DateStyleTable(this.styles);
    const writer = createZipWriter();
    const encoder = new TextEncoder();
    const text = (chunks: () => AsyncIterable<string>) =>
      async function* (): AsyncIterable<Uint8Array> {
        for await (const chunk of chunks()) {
          yield encoder.encode(chunk);
        }
      };

    const edited = [...this.edits].filter(([, edits]) => hasEdits(edits));
    const editedPaths = new Set(edited.map(([index]) => (this.worksheets[index] as WorksheetRef).path));

    // The styles part is declared last on purpose. A date only learns which cell
    // format it needs while its sheet is written, so the part has to be produced
    // after every sheet that might add one.
    for (const entry of this.archive.entries()) {
      if (entry.path === STYLES_PART || editedPaths.has(entry.path)) {
        continue;
      }
      writer.copy(this.archive.storedEntry(entry.path));
    }

    for (const [index, edits] of edited) {
      const ref = this.worksheets[index] as WorksheetRef;
      writer.add(
        ref.path,
        text(() =>
          writeSheetPart(
            this.xml.read(this.archive.openStream(ref.path)),
            {
              positioned: mergeRowEdits(edits.cells, edits.blocks),
              appended: appendedRows(edits.appended),
            },
            { dateStyles, date1904: this.date1904 },
          ),
        ),
      );
    }

    if (this.archive.has(STYLES_PART)) {
      writer.add(
        STYLES_PART,
        text(() => writeStylesPart(this.xml.read(this.archive.openStream(STYLES_PART)), dateStyles)),
      );
    }

    return writer.open();
  }

  private editsFor(index: number): SheetEdits {
    const existing = this.edits.get(index);
    if (existing !== undefined) {
      return existing;
    }

    const created: SheetEdits = { cells: [], blocks: [], appended: [] };
    this.edits.set(index, created);

    return created;
  }
}

class XlsxWorksheetEditor implements WorksheetEditor {
  constructor(
    private readonly edits: SheetEdits,
    private readonly nextOrder: () => number,
  ) {}

  set(ref: string, value: CellInput): this {
    const upper = ref.toUpperCase();
    if (!CELL_REFERENCE.test(upper)) {
      throw new Error(`Not a cell reference: "${ref}"`);
    }

    const row = rowNumberOf(upper);
    const column = columnIndexOf(upper);
    if (row > MAX_ROW || column > MAX_COLUMN_INDEX) {
      throw new Error(`Cell reference out of range: "${ref}"`);
    }

    this.edits.cells.push({ row, column, value, order: this.nextOrder() });

    return this;
  }

  writeRows(startRow: number, rows: RowSource, options: WriteRowsOptions = {}): this {
    if (!Number.isInteger(startRow) || startRow < 1 || startRow > MAX_ROW) {
      throw new Error(`Not a row number: ${startRow}`);
    }

    const { inheritFrom } = options;
    if (inheritFrom !== undefined && (!Number.isInteger(inheritFrom) || inheritFrom < 1)) {
      throw new Error(`Not a row number: ${inheritFrom}`);
    }
    if (inheritFrom !== undefined && inheritFrom > startRow) {
      throw new Error(
        `Cannot inherit formatting from row ${inheritFrom} while writing from row ${startRow}: the sheet is read once from the top, so the row copied from has to come first`,
      );
    }

    this.edits.blocks.push({ startRow, rows, inheritFrom, order: this.nextOrder() });

    return this;
  }

  appendRows(rows: RowSource): this {
    this.edits.appended.push(rows);

    return this;
  }
}

function hasEdits(edits: SheetEdits): boolean {
  return edits.cells.length > 0 || edits.blocks.length > 0 || edits.appended.length > 0;
}

async function* appendedRows(sources: readonly RowSource[]): AsyncIterable<ReadonlyMap<number, CellInput>> {
  for (const source of sources) {
    for await (const values of source) {
      const cells = new Map<number, CellInput>();
      for (const [column, value] of values.entries()) {
        if (value !== undefined) {
          cells.set(column, value);
        }
      }
      yield cells;
    }
  }
}
