import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";

const CONTENT_TYPES_PART = "[Content_Types].xml";

const CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const Element = {
  Override: "Override",
} as const;

const Attribute = {
  PartName: "PartName",
  ContentType: "ContentType",
} as const;

/**
 * Every chart part the workbook holds.
 *
 * Read from the content types rather than by following relationships, because a
 * chart is reached through a drawing and a drawing hangs off either a worksheet or
 * a chartsheet, while the ranges a chart reads can name any sheet at all. The
 * content types name every part in the package, whichever way it is reached.
 */
export async function readChartPaths(archive: ZipArchive, xml: XmlReader): Promise<readonly string[]> {
  if (!archive.has(CONTENT_TYPES_PART)) {
    return [];
  }

  const charts: string[] = [];

  for await (const batch of readPart(archive, xml, CONTENT_TYPES_PART)) {
    for (const event of batch) {
      if (event.type !== "open" || event.name !== Element.Override) {
        continue;
      }
      if (event.attributes[Attribute.ContentType] !== CHART) {
        continue;
      }

      // A content type names a part from the root of the package, with a leading
      // slash an archive path does not have.
      const path = (event.attributes[Attribute.PartName] ?? "").replace(/^\//, "");
      if (archive.has(path)) {
        charts.push(path);
      }
    }
  }

  return charts;
}
