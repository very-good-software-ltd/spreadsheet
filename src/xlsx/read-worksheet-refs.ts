import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readWorkbookRelationships } from "./read-relationships";

const WORKBOOK_PART = "xl/workbook.xml";

const Element = {
  Sheet: "sheet",
} as const;

const Attribute = {
  Name: "name",
  RelationshipId: "r:id",
} as const;

export interface WorksheetRef {
  readonly name: string;
  readonly path: string;
}

export async function readWorksheetRefs(archive: ZipArchive, xml: XmlReader): Promise<WorksheetRef[]> {
  const relationships = await readWorkbookRelationships(archive, xml);
  const refs: WorksheetRef[] = [];

  for await (const event of xml.read(archive.openStream(WORKBOOK_PART))) {
    if (event.type === "open" && event.name === Element.Sheet) {
      const name = event.attributes[Attribute.Name];
      const relId = event.attributes[Attribute.RelationshipId];
      if (name === undefined) {
        continue;
      }
      const path = relId === undefined ? undefined : relationships.get(relId);
      refs.push({ name, path: path ?? "" });
    }
  }

  return refs;
}
