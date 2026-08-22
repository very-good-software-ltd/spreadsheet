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

// The source types the spec allows, mapped to what each one means to us. A
// worksheet range is one we can follow. Consolidation ranges are several, spelled
// in a shape we do not read. External and scenario are somewhere else entirely.
const SOURCE_TYPES: ReadonlyMap<string, PivotSource> = new Map([
  ["worksheet", "worksheet"],
  ["consolidation", "consolidation"],
  ["external", "elsewhere"],
  ["scenario", "elsewhere"],
]);

/**
 * Where a pivot table's cached rows come from, in our terms.
 *
 * An unfamiliar source counts as `consolidation`, the answer that refuses, so a
 * type we have never seen is not quietly assumed to be somewhere we cannot reach.
 */
export type PivotSource = "worksheet" | "consolidation" | "elsewhere";

/** Where a pivot table's cached copy of its source data lives, and what it reads. */
export interface PivotCache {
  readonly path: string;

  readonly source: PivotSource;

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
): Promise<{ source: PivotSource; sheet: string | undefined }> {
  let source: PivotSource = "consolidation";
  let sheet: string | undefined;

  for await (const batch of readPart(archive, xml, path)) {
    for (const event of batch) {
      if (event.type !== "open") {
        continue;
      }
      if (event.name === Element.CacheSource) {
        source = SOURCE_TYPES.get(event.attributes[Attribute.Type] ?? "") ?? "consolidation";
      }
      if (event.name === Element.WorksheetSource) {
        sheet = event.attributes[Attribute.Sheet];
      }
    }
  }

  return { source, sheet };
}
