import { columnIndexOf, rowNumberOf } from "../cell-reference";
import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readRelationships } from "./read-relationships";

const TABLE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

const Element = {
  Table: "table",
} as const;

const Attribute = {
  Name: "name",
  DisplayName: "displayName",
  Extent: "ref",
  HeaderRowCount: "headerRowCount",
  TotalsRowCount: "totalsRowCount",
} as const;

// The attribute is optional and defaults to one, so a table that leaves it out
// still has a header row. Reading it as none would treat the headings as data.
const DEFAULT_HEADER_ROW_COUNT = 1;

export interface TableRef {
  /** The name Excel shows and a formula refers to. */
  readonly name: string;

  /** The table's own part, which has to be rewritten if the table grows. */
  readonly path: string;

  /** The whole extent, header and totals rows included. */
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumnIndex: number;
  readonly lastColumnIndex: number;

  readonly headerRowCount: number;
  readonly totalsRowCount: number;
}

/** Every table on the worksheet at `sheetPath`, in the order it points at them. */
export async function readTables(archive: ZipArchive, xml: XmlReader, sheetPath: string): Promise<readonly TableRef[]> {
  const relationships = await readRelationships(archive, xml, sheetPath);
  const tables: TableRef[] = [];

  for (const relationship of relationships) {
    if (relationship.type !== TABLE_RELATIONSHIP || !archive.has(relationship.target)) {
      continue;
    }

    const table = await readTable(archive, xml, relationship.target);
    if (table !== undefined) {
      tables.push(table);
    }
  }

  return tables;
}

async function readTable(archive: ZipArchive, xml: XmlReader, path: string): Promise<TableRef | undefined> {
  for await (const batch of readPart(archive, xml, path)) {
    for (const event of batch) {
      if (event.type !== "open" || event.name !== Element.Table) {
        continue;
      }

      // A table carries both a name and a displayName, and the display one is what
      // Excel shows and what a formula refers to. Files Excel writes have them
      // agreeing, but the spec lets them differ.
      const name = event.attributes[Attribute.DisplayName] ?? event.attributes[Attribute.Name];
      const extent = event.attributes[Attribute.Extent];
      if (name === undefined || extent === undefined) {
        return undefined;
      }

      return {
        name,
        path,
        ...cornersOf(extent),
        headerRowCount: countOf(event.attributes[Attribute.HeaderRowCount], DEFAULT_HEADER_ROW_COUNT),
        totalsRowCount: countOf(event.attributes[Attribute.TotalsRowCount], 0),
      };
    }
  }

  return undefined;
}

function countOf(attribute: string | undefined, fallback: number): number {
  return attribute === undefined ? fallback : Number(attribute);
}

// A table's extent is written without dollars and without a sheet, since the part
// it lives in already says which sheet it is on.
function cornersOf(extent: string): {
  firstRow: number;
  lastRow: number;
  firstColumnIndex: number;
  lastColumnIndex: number;
} {
  const corners = extent.split(":");
  const rows = corners.map(rowNumberOf);
  const columns = corners.map(columnIndexOf);

  return {
    firstRow: Math.min(...rows),
    lastRow: Math.max(...rows),
    firstColumnIndex: Math.min(...columns),
    lastColumnIndex: Math.max(...columns),
  };
}
