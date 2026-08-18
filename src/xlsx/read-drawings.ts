import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readRelationships } from "./read-relationships";

const DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

/** The drawing parts the worksheet at `sheetPath` places its shapes through. */
export async function readDrawingPaths(
  archive: ZipArchive,
  xml: XmlReader,
  sheetPath: string,
): Promise<readonly string[]> {
  const relationships = await readRelationships(archive, xml, sheetPath);

  return relationships
    .filter((relationship) => relationship.type === DRAWING && archive.has(relationship.target))
    .map((relationship) => relationship.target);
}
