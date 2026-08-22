import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readRelationships } from "./read-relationships";

const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const COMMENTS = `${RELATIONSHIPS}/comments`;
const VML_DRAWING = `${RELATIONSHIPS}/vmlDrawing`;

/** The parts a worksheet's comments are spread across. */
export interface CommentParts {
  /** Where each comment's text and the cell it belongs to are written. */
  readonly comments: readonly string[];

  /**
   * Where the box drawing each comment is positioned. The same parts carry any
   * form controls and any header or footer image the worksheet has.
   */
  readonly vml: readonly string[];
}

/** The comment and VML parts the worksheet at `sheetPath` points at. */
export async function readCommentParts(archive: ZipArchive, xml: XmlReader, sheetPath: string): Promise<CommentParts> {
  const relationships = await readRelationships(archive, xml, sheetPath);
  const targetsOf = (type: string): readonly string[] =>
    relationships
      .filter((relationship) => relationship.type === type && archive.has(relationship.target))
      .map((relationship) => relationship.target);

  return { comments: targetsOf(COMMENTS), vml: targetsOf(VML_DRAWING) };
}
