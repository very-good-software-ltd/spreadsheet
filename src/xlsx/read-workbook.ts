import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readWorkbookRelationships } from "./read-relationships";

const WORKBOOK_PART = "xl/workbook.xml";

const Element = {
  Properties: "workbookPr",
  Sheet: "sheet",
} as const;

const Attribute = {
  Date1904: "date1904",
  Name: "name",
  SheetId: "sheetId",
  RelationshipId: "r:id",
  State: "state",
} as const;

// A sheet's state is "visible", "hidden", or "veryHidden". Only the last two
// keep it out of the tab strip, and veryHidden also hides it from the unhide menu.
const HIDDEN_STATES: ReadonlySet<string> = new Set(["hidden", "veryHidden"]);

export interface WorksheetRef {
  readonly name: string;
  readonly path: string;
  readonly hidden: boolean;

  /** The workbook-wide identifier, distinct from the position and the path. */
  readonly sheetId: number;
}

export interface WorkbookInfo {
  readonly worksheets: readonly WorksheetRef[];
  readonly date1904: boolean;

  /** Every relationship id the part already uses, so a new one can avoid them. */
  readonly relationshipIds: readonly string[];
}

export async function readWorkbook(archive: ZipArchive, xml: XmlReader): Promise<WorkbookInfo> {
  if (!archive.has(WORKBOOK_PART)) {
    throw new Error(`Not a valid xlsx file: missing ${WORKBOOK_PART}`);
  }

  const relationships = await readWorkbookRelationships(archive, xml);
  const worksheets: WorksheetRef[] = [];
  let date1904 = false;

  for await (const batch of readPart(archive, xml, WORKBOOK_PART)) {
    for (const event of batch) {
      if (event.type !== "open") {
        continue;
      }
      if (event.name === Element.Properties) {
        const flag = event.attributes[Attribute.Date1904];
        date1904 = flag === "1" || flag === "true";
      } else if (event.name === Element.Sheet) {
        const name = event.attributes[Attribute.Name];
        if (name === undefined) {
          continue;
        }
        const relId = event.attributes[Attribute.RelationshipId];
        const path = relId === undefined ? undefined : relationships.get(relId);
        const hidden = HIDDEN_STATES.has(event.attributes[Attribute.State] ?? "");
        const sheetId = Number(event.attributes[Attribute.SheetId] ?? "0");
        worksheets.push({ name, path: path ?? "", hidden, sheetId });
      }
    }
  }

  return { worksheets, date1904, relationshipIds: [...relationships.keys()] };
}
