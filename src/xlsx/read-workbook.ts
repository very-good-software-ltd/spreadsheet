import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { type DefinedNameTarget, parseDefinedName } from "./parse-defined-name";
import { readWorkbookRelationships } from "./read-relationships";

const WORKBOOK_PART = "xl/workbook.xml";

const Element = {
  Properties: "workbookPr",
  Sheet: "sheet",
  DefinedName: "definedName",
} as const;

const Attribute = {
  Date1904: "date1904",
  Name: "name",
  SheetId: "sheetId",
  RelationshipId: "r:id",
  State: "state",
  ScopeSheetPosition: "localSheetId",
} as const;

// Excel keeps its own bookkeeping among the defined names, spelled with this
// reserved prefix: the print area, the print titles, the autofilter range. They are
// not names a caller gave anything, so they are not names a caller can address.
const RESERVED_NAME_PREFIX = "_xlnm.";

// A sheet's state is "visible", "hidden", or "veryHidden". Only the last two
// keep it out of the tab strip, and veryHidden also hides it from the unhide menu.
const HIDDEN_STATES: ReadonlySet<string> = new Set(["hidden", "veryHidden"]);

// Not every sheet is a worksheet. A chart drawn on a sheet of its own is a
// chartsheet, and a dialog sheet is another, both listed among the sheets exactly
// like a worksheet and neither holding a row. Which it is shows only in the
// relationship the sheet points through, so one whose relationship names anything
// but a worksheet is left out rather than handed over as one.
const WORKSHEET_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

export interface WorksheetRef {
  readonly name: string;
  readonly path: string;
  readonly hidden: boolean;

  /** The workbook-wide identifier, distinct from the position and the path. */
  readonly sheetId: number;
}

export interface DefinedNameRef {
  readonly name: string;

  /** The sheet the name is scoped to, or `undefined` when it is workbook-wide. */
  readonly scope: string | undefined;

  /** Where the name points, or why it points at nothing writable. */
  readonly target: DefinedNameTarget;
}

export interface WorkbookInfo {
  /** The sheets that hold rows, in document order. A chart sheet is not among them. */
  readonly worksheets: readonly WorksheetRef[];

  /**
   * The largest `sheetId` any sheet carries, chart sheets included, so an added
   * worksheet can take one nothing else already has.
   */
  readonly highestSheetId: number;

  /** Every name the author gave a place in the workbook, in the order they appear. */
  readonly definedNames: readonly DefinedNameRef[];

  readonly date1904: boolean;

  /** Every relationship id the part already uses, so a new one can avoid them. */
  readonly relationshipIds: readonly string[];
}

interface CollectedName {
  readonly name: string;
  readonly scopePosition: number | undefined;
  target: string;
}

interface SheetRef extends WorksheetRef {
  readonly holdsRows: boolean;
}

export async function readWorkbook(archive: ZipArchive, xml: XmlReader): Promise<WorkbookInfo> {
  if (!archive.has(WORKBOOK_PART)) {
    throw new Error(`Not a valid xlsx file: missing ${WORKBOOK_PART}`);
  }

  const relationships = await readWorkbookRelationships(archive, xml);
  const sheets: SheetRef[] = [];
  const collected: CollectedName[] = [];
  let collecting: CollectedName | undefined;
  let date1904 = false;

  for await (const batch of readPart(archive, xml, WORKBOOK_PART)) {
    for (const event of batch) {
      if (event.type === "text") {
        if (collecting !== undefined) {
          collecting.target += event.text;
        }
        continue;
      }
      if (event.type === "close") {
        if (event.name === Element.DefinedName && collecting !== undefined) {
          collected.push(collecting);
          collecting = undefined;
        }
        continue;
      }
      if (event.name === Element.DefinedName) {
        const name = event.attributes[Attribute.Name];
        if (name === undefined || name.startsWith(RESERVED_NAME_PREFIX)) {
          continue;
        }
        const scopePosition = event.attributes[Attribute.ScopeSheetPosition];
        collecting = {
          name,
          scopePosition: scopePosition === undefined ? undefined : Number(scopePosition),
          target: "",
        };
      } else if (event.name === Element.Properties) {
        const flag = event.attributes[Attribute.Date1904];
        date1904 = flag === "1" || flag === "true";
      } else if (event.name === Element.Sheet) {
        const name = event.attributes[Attribute.Name];
        if (name === undefined) {
          continue;
        }
        const relId = event.attributes[Attribute.RelationshipId];
        const relationship = relId === undefined ? undefined : relationships.get(relId);
        const hidden = HIDDEN_STATES.has(event.attributes[Attribute.State] ?? "");
        const sheetId = Number(event.attributes[Attribute.SheetId] ?? "0");
        // Only a relationship that positively says the sheet is something else
        // drops it. One that is missing, or that names no type, has said nothing,
        // and a sheet we cannot place is better failing when its rows are asked for
        // than disappearing in silence.
        const relationshipType = relationship?.type ?? "";
        const holdsRows = relationshipType === "" || relationshipType === WORKSHEET_RELATIONSHIP;
        sheets.push({ name, path: relationship?.target ?? "", hidden, sheetId, holdsRows });
      }
    }
  }

  // A name's scope is given as the sheet's zero-based position, not as the `sheetId`
  // the sheet carries, so it resolves against the sheet order rather than that id.
  // The two disagree in real files: in `benchmark/files/Electricity_since_1920.xlsx`,
  // `localSheetId="3"` means `Fuel Input`, the fourth sheet, whose `sheetId` is 12.
  const definedNames = collected.map((defined) => ({
    name: defined.name,
    scope: defined.scopePosition === undefined ? undefined : sheets[defined.scopePosition]?.name,
    target: parseDefinedName(defined.target),
  }));

  return {
    worksheets: sheets.filter((sheet) => sheet.holdsRows),
    highestSheetId: Math.max(0, ...sheets.map((sheet) => sheet.sheetId)),
    definedNames,
    date1904,
    relationshipIds: [...relationships.keys()],
  };
}
