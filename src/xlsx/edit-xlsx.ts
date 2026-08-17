import type { CellInput } from "../cell-input";
import { columnIndexOf, rowNumberOf } from "../cell-reference";
import type { Editor, RowSource, WorksheetEditor, WriteRowsOptions } from "../editor";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlEvent, XmlReader } from "../xml/xml-reader";
import { createZipWriter } from "../zip/create-zip-writer";
import type { ZipArchive } from "../zip/zip-archive";
import { MAIN_NAMESPACE } from "./blank-workbook";
import { type CellEdit, mergeRowEdits, type RowBlock } from "./merge-row-edits";
import type { Styles } from "./read-styles";
import type { WorkbookInfo } from "./read-workbook";
import {
  type AddedWorksheet,
  asPart,
  flatten,
  withAddedContentTypes,
  withAddedRelationships,
  withAddedSheets,
  withoutContentTypeOverride,
  withoutRelationshipTo,
  withRecalculationOnLoad,
} from "./write-package";
import { writeSheetPart } from "./write-sheet";
import { DateStyleTable, writeStylesPart } from "./write-styles";

const STYLES_PART = "xl/styles.xml";
const WORKBOOK_PART = "xl/workbook.xml";
const WORKBOOK_RELATIONSHIPS_PART = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES_PART = "[Content_Types].xml";

// The calculation chain records the order a spreadsheet application evaluated the
// formulas in. It is a cache, rebuilt on open, and an edit invalidates it, so it
// is left out of the written file rather than carried across wrong. Its content
// type override and its relationship go with it, since the part they name is no
// longer there.
const CALCULATION_CHAIN_PART = "xl/calcChain.xml";
const CALCULATION_CHAIN_TARGET = "calcChain.xml";

const CELL_REFERENCE = /^[A-Z]{1,3}[1-9][0-9]*$/;

// The last row and column a worksheet can hold, and the longest a name can be.
const MAX_ROW = 1_048_576;
const MAX_COLUMN_INDEX = 16_383;
const MAX_WORKSHEET_NAME_LENGTH = 31;

// A formula refers to a sheet by name, wrapping it in single quotes when it needs
// to, so none of these can appear in one.
const FORBIDDEN_IN_WORKSHEET_NAME = /[:\\/?*[\]]/;

interface SheetEdits {
  readonly cells: CellEdit[];
  readonly blocks: RowBlock[];
  readonly appended: RowSource[];
}

interface Target {
  readonly name: string;
  readonly path: string;
  readonly added: boolean;
}

export class XlsxEditor implements Editor {
  private readonly targets: Target[];
  private readonly edits = new Map<number, SheetEdits>();
  private readonly added: AddedWorksheet[] = [];
  private readonly takenRelationshipIds: Set<string>;
  private readonly xml: XmlReader = createXmlReader();
  private nextSheetId: number;
  private calls = 0;
  private saved = false;

  constructor(
    private readonly archive: ZipArchive,
    private readonly workbook: WorkbookInfo,
    private readonly styles: Styles,
  ) {
    this.targets = workbook.worksheets.map((sheet) => ({ name: sheet.name, path: sheet.path, added: false }));
    this.takenRelationshipIds = new Set(workbook.relationshipIds);
    this.nextSheetId = Math.max(0, ...workbook.worksheets.map((sheet) => sheet.sheetId)) + 1;
  }

  worksheet(nameOrIndex: string | number): WorksheetEditor {
    const index =
      typeof nameOrIndex === "number" ? nameOrIndex : this.targets.findIndex((sheet) => sheet.name === nameOrIndex);

    if (index < 0 || index >= this.targets.length) {
      throw new Error(`Worksheet not found: ${nameOrIndex}`);
    }

    return this.editorFor(index);
  }

  addWorksheet(name: string): WorksheetEditor {
    this.checkWorksheetName(name);

    const sheet: AddedWorksheet = {
      name,
      path: this.freeWorksheetPath(),
      relationshipId: this.freeRelationshipId(),
      sheetId: this.nextSheetId,
    };
    this.nextSheetId += 1;
    this.added.push(sheet);
    this.targets.push({ name, path: sheet.path, added: true });

    return this.editorFor(this.targets.length - 1);
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
    const part = (path: string): AsyncIterable<XmlEvent> => flatten(this.xml.read(this.archive.openStream(path)));

    const rewritten = this.packageParts(part);
    const sheets = this.sheetParts(dateStyles);

    for (const entry of this.archive.entries()) {
      if (entry.path === STYLES_PART || sheets.has(entry.path) || entry.path === CALCULATION_CHAIN_PART) {
        continue;
      }

      const rewrite = rewritten.get(entry.path);
      if (rewrite === undefined) {
        writer.copy(this.archive.storedEntry(entry.path));
      } else {
        writer.add(entry.path, encoded(rewrite));
      }
    }

    for (const [path, content] of sheets) {
      writer.add(path, encoded(content));
    }

    // The styles part is declared last on purpose. A date only learns which cell
    // format it needs while its sheet is written, so the part has to be produced
    // after every sheet that might add one.
    if (this.archive.has(STYLES_PART)) {
      writer.add(
        STYLES_PART,
        encoded(() => writeStylesPart(this.xml.read(this.archive.openStream(STYLES_PART)), dateStyles)),
      );
    }

    return writer.open();
  }

  // The parts that describe the package rather than hold data. Each is rewritten
  // only when something about it has to change, so the rest stay byte-identical.
  private packageParts(part: (path: string) => AsyncIterable<XmlEvent>): Map<string, () => AsyncIterable<string>> {
    const droppingChain = this.archive.has(CALCULATION_CHAIN_PART);
    const rewritten = new Map<string, () => AsyncIterable<string>>();

    rewritten.set(WORKBOOK_PART, () =>
      asPart(withAddedSheets(withRecalculationOnLoad(part(WORKBOOK_PART)), this.added)),
    );

    if (!droppingChain && this.added.length === 0) {
      return rewritten;
    }

    rewritten.set(CONTENT_TYPES_PART, () => {
      const source = part(CONTENT_TYPES_PART);
      const trimmed = droppingChain ? withoutContentTypeOverride(source, `/${CALCULATION_CHAIN_PART}`) : source;
      return asPart(withAddedContentTypes(trimmed, this.added));
    });

    rewritten.set(WORKBOOK_RELATIONSHIPS_PART, () => {
      const source = part(WORKBOOK_RELATIONSHIPS_PART);
      const trimmed = droppingChain ? withoutRelationshipTo(source, CALCULATION_CHAIN_TARGET) : source;
      return asPart(withAddedRelationships(trimmed, this.added));
    });

    return rewritten;
  }

  private sheetParts(dateStyles: DateStyleTable): Map<string, () => AsyncIterable<string>> {
    const parts = new Map<string, () => AsyncIterable<string>>();

    for (const [index, target] of this.targets.entries()) {
      const edits = this.edits.get(index);

      if (!target.added && (edits === undefined || !hasEdits(edits))) {
        continue;
      }

      parts.set(target.path, () =>
        writeSheetPart(
          target.added ? emptyWorksheet() : this.xml.read(this.archive.openStream(target.path)),
          {
            positioned: mergeRowEdits(edits?.cells ?? [], edits?.blocks ?? []),
            appended: appendedRows(edits?.appended ?? []),
            inheritedRows: inheritedRows(edits?.blocks ?? []),
          },
          { dateStyles, date1904: this.workbook.date1904 },
        ),
      );
    }

    return parts;
  }

  private editorFor(index: number): WorksheetEditor {
    const edits = this.edits.get(index) ?? { cells: [], blocks: [], appended: [] };
    this.edits.set(index, edits);

    return new XlsxWorksheetEditor(edits, () => {
      this.calls += 1;
      return this.calls;
    });
  }

  private checkWorksheetName(name: string): void {
    if (name === "" || name.length > MAX_WORKSHEET_NAME_LENGTH) {
      throw new Error(`Not a worksheet name: "${name}". A name is 1 to ${MAX_WORKSHEET_NAME_LENGTH} characters`);
    }
    if (FORBIDDEN_IN_WORKSHEET_NAME.test(name) || name.startsWith("'") || name.endsWith("'")) {
      throw new Error(
        `Not a worksheet name: "${name}". A formula refers to a sheet by name, so : \\ / ? * [ ] and a surrounding quote cannot appear in one`,
      );
    }
    if (this.targets.some((sheet) => sheet.name.toLowerCase() === name.toLowerCase())) {
      throw new Error(`Worksheet already exists: "${name}"`);
    }
  }

  private freeWorksheetPath(): string {
    for (let number = 1; ; number += 1) {
      const path = `xl/worksheets/sheet${number}.xml`;
      if (!this.archive.has(path) && !this.targets.some((sheet) => sheet.path === path)) {
        return path;
      }
    }
  }

  private freeRelationshipId(): string {
    for (let number = 1; ; number += 1) {
      const id = `rId${number}`;
      if (!this.takenRelationshipIds.has(id)) {
        this.takenRelationshipIds.add(id);
        return id;
      }
    }
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

// Every block declares which row it copies formatting from up front, even though
// its rows are only read at save, so the sheet reader knows what to hold on to.
function inheritedRows(blocks: readonly RowBlock[]): ReadonlySet<number> {
  return new Set(blocks.flatMap((block) => (block.inheritFrom === undefined ? [] : [block.inheritFrom])));
}

function hasEdits(edits: SheetEdits): boolean {
  return edits.cells.length > 0 || edits.blocks.length > 0 || edits.appended.length > 0;
}

function encoded(chunks: () => AsyncIterable<string>): () => AsyncIterable<Uint8Array> {
  const encoder = new TextEncoder();

  return async function* (): AsyncIterable<Uint8Array> {
    for await (const chunk of chunks()) {
      yield encoder.encode(chunk);
    }
  };
}

// A sheet being added has no source part to transform, so it starts as the events
// an empty one would have produced and goes through the same writer.
async function* emptyWorksheet(): AsyncIterable<readonly XmlEvent[]> {
  yield [
    { type: "open", name: "worksheet", attributes: { xmlns: MAIN_NAMESPACE } },
    { type: "open", name: "sheetData", attributes: {} },
    { type: "close", name: "sheetData" },
    { type: "close", name: "worksheet" },
  ];
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
