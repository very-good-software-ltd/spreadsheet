import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readPart } from "./read-part";

const RELATIONSHIPS_PART = "xl/_rels/workbook.xml.rels";

const Element = {
  Relationship: "Relationship",
} as const;

const Attribute = {
  Id: "Id",
  Target: "Target",
} as const;

export async function readWorkbookRelationships(archive: ZipArchive, xml: XmlReader): Promise<Map<string, string>> {
  const targets = new Map<string, string>();
  if (!archive.has(RELATIONSHIPS_PART)) {
    return targets;
  }

  for await (const event of readPart(archive, xml, RELATIONSHIPS_PART)) {
    if (event.type === "open" && event.name === Element.Relationship) {
      const id = event.attributes[Attribute.Id];
      const target = event.attributes[Attribute.Target];
      if (id !== undefined && target !== undefined) {
        targets.set(id, resolvePartPath(target));
      }
    }
  }

  return targets;
}

function resolvePartPath(target: string): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  return `xl/${target}`;
}
