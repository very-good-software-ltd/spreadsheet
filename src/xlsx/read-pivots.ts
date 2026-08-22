import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readRelationships } from "./read-relationships";

const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PIVOT_CACHE = `${RELATIONSHIPS}/pivotCacheDefinition`;
const PIVOT_TABLE = `${RELATIONSHIPS}/pivotTable`;

const WORKBOOK_PART = "xl/workbook.xml";

const Element = {
  CacheSource: "cacheSource",
  WorksheetSource: "worksheetSource",
} as const;

const Attribute = {
  Type: "type",
  Sheet: "sheet",
} as const;

// The one source type whose data lives in this workbook. The others read a cube, an
// external query or a scenario, none of which a region can move.
const FROM_WORKSHEET = "worksheet";

/** Where a pivot table's cached copy of its source data lives, and what it reads. */
export interface PivotCache {
  readonly path: string;

  /** Whether the cached rows come from a range in this workbook. */
  readonly fromWorksheet: boolean;

  /**
   * The worksheet the range is on. Absent when the source is given as a name rather
   * than spelled out, in which case the name says which sheet and we do not resolve
   * it.
   */
  readonly sheet: string | undefined;
}

/** Every pivot cache the workbook holds, since a cache belongs to it and not to a sheet. */
export async function readPivotCaches(archive: ZipArchive, xml: XmlReader): Promise<readonly PivotCache[]> {
  const caches: PivotCache[] = [];

  for (const relationship of await readRelationships(archive, xml, WORKBOOK_PART)) {
    if (relationship.type !== PIVOT_CACHE || !archive.has(relationship.target)) {
      continue;
    }

    caches.push({ path: relationship.target, ...(await sourceOf(archive, xml, relationship.target)) });
  }

  return caches;
}

/** The pivot tables drawn on the worksheet at `sheetPath`. */
export async function readPivotTablePaths(
  archive: ZipArchive,
  xml: XmlReader,
  sheetPath: string,
): Promise<readonly string[]> {
  const relationships = await readRelationships(archive, xml, sheetPath);

  return relationships
    .filter((relationship) => relationship.type === PIVOT_TABLE && archive.has(relationship.target))
    .map((relationship) => relationship.target);
}

async function sourceOf(
  archive: ZipArchive,
  xml: XmlReader,
  path: string,
): Promise<{ fromWorksheet: boolean; sheet: string | undefined }> {
  let fromWorksheet = false;
  let sheet: string | undefined;

  for await (const batch of readPart(archive, xml, path)) {
    for (const event of batch) {
      if (event.type !== "open") {
        continue;
      }
      if (event.name === Element.CacheSource) {
        fromWorksheet = event.attributes[Attribute.Type] === FROM_WORKSHEET;
      }
      if (event.name === Element.WorksheetSource) {
        sheet = event.attributes[Attribute.Sheet];
      }
    }
  }

  return { fromWorksheet, sheet };
}
