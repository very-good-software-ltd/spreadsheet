import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readRelationships } from "./read-relationships";
import type { RowShift } from "./shift-formula";

const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PIVOT_CACHE = `${RELATIONSHIPS}/pivotCacheDefinition`;

const WORKBOOK_PART = "xl/workbook.xml";

const Element = {
  PivotSource: "worksheetSource",
} as const;

const Attribute = {
  Reference: "ref",
  Sheet: "sheet",
} as const;

/**
 * Everything on the worksheet named `sheetName` that sits where `shift` would move
 * rows and that we cannot move with it, each described so a caller can say which.
 *
 * Empty when the sheet can be moved, which is the ordinary case for a template of
 * headings, data, totals and formatting.
 */
export async function blockersFor(
  archive: ZipArchive,
  xml: XmlReader,
  sheetName: string,
  shift: RowShift,
): Promise<readonly string[]> {
  const blockers: string[] = [];

  if (await pivotReaches(archive, xml, sheetName, shift.at)) {
    blockers.push(`a pivot table reading from row ${shift.at} or below, whose cached source range we do not rewrite`);
  }

  return blockers;
}

// A pivot cache belongs to the workbook rather than to a sheet, and names the sheet
// it reads from, so it is looked at whichever sheet is moving.
async function pivotReaches(archive: ZipArchive, xml: XmlReader, sheet: string, row: number): Promise<boolean> {
  for (const relationship of await readRelationships(archive, xml, WORKBOOK_PART)) {
    if (relationship.type !== PIVOT_CACHE || !archive.has(relationship.target)) {
      continue;
    }

    for await (const batch of readPart(archive, xml, relationship.target)) {
      for (const event of batch) {
        if (event.type !== "open" || event.name !== Element.PivotSource) {
          continue;
        }
        if (event.attributes[Attribute.Sheet] === sheet && lastRowOf(event.attributes[Attribute.Reference]) >= row) {
          return true;
        }
      }
    }
  }

  return false;
}

function rowOf(reference: string | undefined): number {
  return Number((reference ?? "").replaceAll(/[$A-Z]/g, ""));
}

function lastRowOf(extent: string | undefined): number {
  return rowOf((extent ?? "").split(":").at(-1));
}
