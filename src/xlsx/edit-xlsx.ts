import type { CellInput } from "../cell-input";
import { columnIndexOf, rowNumberOf } from "../cell-reference";
import type { Editor, RowSource, WorksheetEditor, WriteRowsOptions } from "../editor";
import type { NamedRegion } from "../named-region";
import { placeEach, readRegionRows } from "../region-rows";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlEvent, XmlReader } from "../xml/xml-reader";
import { createZipWriter } from "../zip/create-zip-writer";
import type { ZipArchive } from "../zip/zip-archive";
import { MAIN_NAMESPACE } from "./blank-workbook";
import { type CellEdit, mergeRowEdits, type RowBlock } from "./merge-row-edits";
import { blockersFor } from "./movement-blockers";
import type { CommentParts } from "./read-comments";
import type { Styles } from "./read-styles";
import type { TableOnSheet } from "./read-tables";
import type { WorkbookInfo } from "./read-workbook";
import { shiftFor } from "./region-shift";
import { type NamedThings, resolveRegion } from "./resolve-region";
import { shiftCommentRefs } from "./shift-comment";
import { shiftDrawingAnchors } from "./shift-drawing";
import type { RowShift } from "./shift-formula";
import { shiftForeignFormulas } from "./shift-sheet";
import { shiftVmlAnchors } from "./shift-vml";
import {
  type AddedWorksheet,
  asPart,
  flatten,
  withAddedContentTypes,
  withAddedRelationships,
  withAddedSheets,
  withMovedDefinedNames,
  withoutContentTypeOverride,
  withoutRelationshipTo,
  withRecalculationOnLoad,
} from "./write-package";
import { type RegionWrite, writeSheetPart } from "./write-sheet";
import { DateStyleTable, writeStylesPart } from "./write-styles";
import { withTableExtent } from "./write-table";

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

interface PreparedRegion {
  readonly sheet: string;
  readonly region: NamedRegion;
  readonly rows: AsyncIterable<readonly (CellInput | undefined)[]>;

  /** Settled once the worksheet holding it has been written and its rows counted. */
  readonly shift: Promise<RowShift | undefined>;
  readonly moved: (shift: RowShift | undefined) => void;
}

interface RegionEdit {
  readonly region: NamedRegion;
  readonly rows: RowSource;
  readonly table: TableOnSheet | undefined;
  readonly order: number;
}

interface SheetEdits {
  readonly cells: CellEdit[];
  readonly blocks: RowBlock[];
  readonly appended: RowSource[];
  readonly regions: RegionEdit[];
}

interface Target {
  readonly name: string;
  readonly path: string;
  readonly added: boolean;
}

/** A part positioned by row, and how to move it once its worksheet has moved. */
interface AnchoredPart {
  readonly sheet: string;
  readonly shift: (events: AsyncIterable<readonly XmlEvent[]>, shift: RowShift) => AsyncIterable<readonly XmlEvent[]>;
}

export class XlsxEditor implements Editor {
  private readonly targets: Target[];
  private readonly edits = new Map<number, SheetEdits>();
  private readonly added: AddedWorksheet[] = [];
  private readonly growing = new Map<string, { lastRow: number }>();

  // Every region is read before any part is written, so that each worksheet can be
  // given its own move and every other worksheet's as well. A formula can name any
  // sheet, so no order of writing would let one region learn about a later one.
  private prepared: Promise<readonly PreparedRegion[]> | undefined;
  private readonly takenRelationshipIds: Set<string>;
  private readonly xml: XmlReader = createXmlReader();
  private nextSheetId: number;
  private calls = 0;
  private saved = false;

  constructor(
    private readonly archive: ZipArchive,
    private readonly workbook: WorkbookInfo,
    private readonly styles: Styles,
    private readonly tables: readonly TableOnSheet[] = [],
    private readonly drawings: ReadonlyMap<string, readonly string[]> = new Map(),
    private readonly comments: ReadonlyMap<string, CommentParts> = new Map(),
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

  writeRegion(name: string, rows: RowSource): this {
    const { region, table } = resolveRegion(this.named, name);
    const index = this.targets.findIndex((sheet) => sheet.name === region.sheet);

    if (index < 0) {
      throw new Error(
        `The name "${region.name}" points at worksheet "${region.sheet}", which this workbook has not got`,
      );
    }

    this.editorFor(index);
    this.regionEditsFor(index).push({ region, rows, table, order: this.nextOrder() });
    this.registerGrowth(table);

    return this;
  }

  // The table's part is rewritten at save with where it ended up, so it has to be
  // kept out of the entries copied across, and that choice is made before any row
  // has been read.
  private registerGrowth(table: TableOnSheet | undefined): void {
    if (table !== undefined && !this.growing.has(table.path)) {
      this.growing.set(table.path, { lastRow: table.lastRow });
    }
  }

  private regionEditsFor(index: number): RegionEdit[] {
    const edits = this.edits.get(index);
    if (edits === undefined) {
      throw new Error(`No edits for worksheet ${index}`);
    }
    return edits.regions;
  }

  private nextOrder(): number {
    this.calls += 1;
    return this.calls;
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
    const anchored = this.anchoredParts();

    for (const entry of this.archive.entries()) {
      if (
        entry.path === STYLES_PART ||
        entry.path === WORKBOOK_PART ||
        sheets.has(entry.path) ||
        entry.path === CALCULATION_CHAIN_PART ||
        this.growing.has(entry.path) ||
        anchored.has(entry.path)
      ) {
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

    // Everything a moved sheet positions by row is declared after it, because how
    // far its rows moved is only known once they have gone past.
    for (const [path, moving] of anchored) {
      writer.add(
        path,
        encoded(() => this.anchoredPart(path, moving)),
      );
    }

    // The workbook part is declared after the sheets because its defined names have
    // to move with the rows, and how far they moved is only known once the sheet
    // holding the region has been written.
    if (this.archive.has(WORKBOOK_PART)) {
      const workbook = rewritten.get(WORKBOOK_PART);
      if (workbook !== undefined) {
        writer.add(WORKBOOK_PART, encoded(workbook));
      }
    }

    // A grown table's part is declared after the sheets for the same reason the
    // styles part is: how far the table reached is only known once the rows that
    // filled it have gone past.
    for (const [path, growth] of this.growing) {
      writer.add(
        path,
        encoded(() => asPart(withTableExtent(part(path), growth.lastRow))),
      );
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
      asPart(withAddedSheets(withRecalculationOnLoad(this.workbookNames(part(WORKBOOK_PART))), this.added)),
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

    // A region moves rows, and any worksheet's formulas can name them, so once one
    // is written every worksheet is rebuilt rather than copied.
    const moving = [...this.edits.values()].some((edits) => edits.regions.length > 0);

    // A sheet holding a region goes out before the rest, whatever order the workbook
    // lists them in. It is the only one that learns how far its rows moved, and it
    // learns it by writing them, so a sheet written before it would wait forever.
    for (const [index, target] of this.regionFirst()) {
      const edits = this.edits.get(index);

      if (!target.added && !moving && (edits === undefined || !hasEdits(edits))) {
        continue;
      }

      parts.set(target.path, () => this.sheetPart(target, edits, dateStyles));
    }

    return parts;
  }

  private regionFirst(): readonly (readonly [number, Target])[] {
    const ordered = [...this.targets.entries()].map(([index, target]) => [index, target] as const);

    return ordered.sort(([left], [right]) => regionRank(this.edits.get(left)) - regionRank(this.edits.get(right)));
  }

  // A region is the one edit that moves the rest of the sheet, so its rows are read
  // before anything is written. How far the sheet moves depends on how many there
  // are, and the rows above the region go out before the ones inside it.
  private async *sheetPart(
    target: Target,
    edits: SheetEdits | undefined,
    dateStyles: DateStyleTable,
  ): AsyncIterable<string> {
    const regions = await this.regions();
    const own = regions.find((prepared) => prepared.sheet === target.name);

    let events: AsyncIterable<readonly XmlEvent[]> = target.added
      ? emptyWorksheet()
      : this.xml.read(this.archive.openStream(target.path));

    if (own !== undefined) {
      await this.refuseIfBlocked(target, own);
    }

    // Its own move is applied as the sheet is written, because only the writer
    // knows how far it goes. Every other move is already settled by then, and only
    // a qualified reference can name the sheet it happened on.
    for (const other of regions) {
      if (other !== own) {
        const shift = await other.shift;
        if (shift !== undefined) {
          events = shiftForeignFormulas(events, shift, target.name);
        }
      }
    }

    yield* writeSheetPart(
      events,
      {
        positioned: mergeRowEdits(edits?.cells ?? [], edits?.blocks ?? []),
        appended: appendedRows(edits?.appended ?? []),
        inheritedRows: inheritedRows(edits?.blocks ?? []),
        ...(own === undefined ? {} : { region: this.regionWrite(own, edits) }),
      },
      { dateStyles, date1904: this.workbook.date1904 },
    );
  }

  private regionWrite(prepared: PreparedRegion, edits: SheetEdits | undefined): RegionWrite {
    const table = edits?.regions[0]?.table;

    return {
      firstRow: prepared.region.firstRow,
      lastRow: prepared.region.lastRow,
      rows: prepared.rows,
      moveFor: (count) => shiftFor(prepared.region, count),
      moved: (shift) => {
        this.growTable(table, shift);
        prepared.moved(shift);
      },
    };
  }

  // Which parts positioned by row belong to a worksheet a region is being written
  // into. That a region was written is known before any row is read, which is when
  // the choice to copy an entry or rebuild it has to be made.
  private anchoredParts(): Map<string, AnchoredPart> {
    const moving = new Map<string, AnchoredPart>();

    for (const [index, edits] of this.edits) {
      const name = this.targets[index]?.name;
      if (name === undefined || edits.regions.length === 0) {
        continue;
      }

      const comments = this.comments.get(name);

      for (const path of this.drawings.get(name) ?? []) {
        moving.set(path, { sheet: name, shift: shiftDrawingAnchors });
      }
      for (const path of comments?.comments ?? []) {
        moving.set(path, { sheet: name, shift: shiftCommentRefs });
      }
      for (const path of comments?.vml ?? []) {
        moving.set(path, { sheet: name, shift: shiftVmlAnchors });
      }
    }

    return moving;
  }

  private async *anchoredPart(path: string, moving: AnchoredPart): AsyncIterable<string> {
    const shift = await (await this.regions()).find((region) => region.sheet === moving.sheet)?.shift;
    const events = this.xml.read(this.archive.openStream(path));

    yield* asPart(flatten(shift === undefined ? events : moving.shift(events, shift)));
  }

  // Asked before the rows are counted, so the row the move starts at is not known
  // yet. The region's own first row is the earliest it can be, which refuses a
  // little more than it has to rather than a little less.
  private async refuseIfBlocked(target: Target, prepared: PreparedRegion): Promise<void> {
    const at = prepared.region.firstRow;
    const blockers = await blockersFor(this.archive, this.xml, target.name, { sheet: prepared.sheet, at, by: 0 });

    if (blockers.length > 0) {
      throw new Error(
        `Cannot write into "${prepared.region.name}": it moves the rows of worksheet "${target.name}" from row ${at}, and that sheet has ${blockers.join(", and ")}`,
      );
    }
  }

  private regions(): Promise<readonly PreparedRegion[]> {
    this.prepared ??= this.prepareRegions();

    return this.prepared;
  }

  private async prepareRegions(): Promise<readonly PreparedRegion[]> {
    const written = [...this.edits].flatMap(([index, edits]) => {
      const region = firstRegion(edits, this.targets[index]?.name ?? "");
      return region === undefined ? [] : [region];
    });

    const prepared: PreparedRegion[] = [];

    for (const edit of written) {
      const settled = settlement<RowShift | undefined>();

      // With one region its rows are counted as they are written, so nothing is
      // held. With more than one, a formula on either sheet can name the other, and
      // no order of writing lets both learn about the other first, so both are
      // counted up front instead.
      if (written.length === 1) {
        prepared.push({
          sheet: edit.region.sheet,
          region: edit.region,
          rows: placeEach(edit.region, edit.rows),
          shift: settled.value,
          moved: settled.settle,
        });
        continue;
      }

      const rows = await readRegionRows(edit.region, edit.rows);
      const shift = shiftFor(edit.region, rows.count);
      settled.settle(shift);
      this.growTable(edit.table, shift);

      prepared.push({
        sheet: edit.region.sheet,
        region: edit.region,
        rows: rows.rows(),
        shift: settled.value,
        moved: settled.settle,
      });
    }

    return prepared;
  }

  private growTable(table: TableOnSheet | undefined, shift: RowShift | undefined): void {
    if (table === undefined || shift === undefined) {
      return;
    }

    const growth = this.growing.get(table.path);
    if (growth !== undefined) {
      growth.lastRow = table.lastRow + shift.by;
    }
  }

  // Every name the moves touch, moved with them, so a name still covers what its
  // author drew around and the written file can be filled again.
  private async *workbookNames(events: AsyncIterable<XmlEvent>): AsyncIterable<XmlEvent> {
    let moved = events;

    for (const region of await this.regions()) {
      const shift = await region.shift;
      if (shift !== undefined) {
        moved = withMovedDefinedNames(moved, shift);
      }
    }

    yield* moved;
  }

  private editorFor(index: number): WorksheetEditor {
    const edits = this.edits.get(index) ?? { cells: [], blocks: [], appended: [], regions: [] };
    this.edits.set(index, edits);

    return new XlsxWorksheetEditor(
      edits,
      this.named,
      this.targets[index]?.name ?? "",
      (table) => this.registerGrowth(table),
      () => this.nextOrder(),
    );
  }

  private get named(): NamedThings {
    return { definedNames: this.workbook.definedNames, tables: this.tables };
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
    private readonly named: NamedThings,
    private readonly sheetName: string,
    private readonly registerGrowth: (table: TableOnSheet | undefined) => void,
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

  writeRegion(name: string, rows: RowSource): this {
    const { region, table } = resolveRegion(this.named, name, this.sheetName);

    this.edits.regions.push({ region, rows, table, order: this.nextOrder() });
    this.registerGrowth(table);

    return this;
  }
}

// Only one region can be written per worksheet. Two would each move the sheet under
// the other, and the second's rows would be aimed at where the first used to be.
function firstRegion(edits: SheetEdits | undefined, sheet: string): RegionEdit | undefined {
  const regions = edits?.regions ?? [];

  if (regions.length > 1) {
    throw new Error(
      `Worksheet "${sheet}" has ${regions.length} regions written to it. Writing one moves the rows the others were aimed at, so only one region per worksheet can be written`,
    );
  }

  return regions[0];
}

// Every block declares which row it copies formatting from up front, even though
// its rows are only read at save, so the sheet reader knows what to hold on to.
function inheritedRows(blocks: readonly RowBlock[]): ReadonlySet<number> {
  return new Set(blocks.flatMap((block) => (block.inheritFrom === undefined ? [] : [block.inheritFrom])));
}

function regionRank(edits: SheetEdits | undefined): number {
  return (edits?.regions.length ?? 0) > 0 ? 0 : 1;
}

function hasEdits(edits: SheetEdits): boolean {
  return edits.cells.length > 0 || edits.blocks.length > 0 || edits.appended.length > 0 || edits.regions.length > 0;
}

// A part is produced a row at a time, and each row is small. Encoding and handing
// over every one of them separately means a million trips through the encoder and
// a million awaited writes into the compressor for a million-row sheet, which
// costs more than the work itself. Gathering them into blocks first crosses those
// boundaries once per block instead. This is the same cost the reader hit yielding
// one XML event at a time, and the same fix.
const OUTPUT_BLOCK_CHARACTERS = 64 * 1024;

function encoded(chunks: () => AsyncIterable<string>): () => AsyncIterable<Uint8Array> {
  const encoder = new TextEncoder();

  return async function* (): AsyncIterable<Uint8Array> {
    let held = "";

    for await (const chunk of chunks()) {
      held += chunk;

      if (held.length >= OUTPUT_BLOCK_CHARACTERS) {
        yield encoder.encode(held);
        held = "";
      }
    }

    if (held !== "") {
      yield encoder.encode(held);
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

// A value that is not known when it is asked for, only when the work that produces
// it is done. Used for how far a sheet's rows moved, which its own writer settles
// and every other part of the file waits on.
function settlement<T>(): { value: Promise<T>; settle: (value: T) => void } {
  let settle: ((value: T) => void) | undefined;
  const value = new Promise<T>((resolve) => {
    settle = resolve;
  });

  return { value, settle: (settled) => settle?.(settled) };
}
