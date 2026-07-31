import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readWorkbookRelationships } from "./read-relationships";

const WORKBOOK_PART = "xl/workbook.xml";

export interface WorksheetRef {
  readonly name: string;
  readonly path: string;
}

export async function readWorksheetRefs(archive: ZipArchive, xml: XmlReader): Promise<WorksheetRef[]> {
  const relationships = await readWorkbookRelationships(archive, xml);
  const refs: WorksheetRef[] = [];

  for await (const event of xml.read(archive.openStream(WORKBOOK_PART))) {
    if (event.type === "open" && event.name === "sheet") {
      const name = event.attributes["name"];
      const relId = event.attributes["r:id"];
      if (name === undefined) {
        continue;
      }
      const path = relId === undefined ? undefined : relationships.get(relId);
      refs.push({ name, path: path ?? "" });
    }
  }

  return refs;
}
