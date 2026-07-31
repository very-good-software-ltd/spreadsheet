import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";

const RELATIONSHIPS_PART = "xl/_rels/workbook.xml.rels";

export async function readWorkbookRelationships(archive: ZipArchive, xml: XmlReader): Promise<Map<string, string>> {
  const targets = new Map<string, string>();
  if (!archive.has(RELATIONSHIPS_PART)) {
    return targets;
  }

  for await (const event of xml.read(archive.openStream(RELATIONSHIPS_PART))) {
    if (event.type === "open" && event.name === "Relationship") {
      const id = event.attributes["Id"];
      const target = event.attributes["Target"];
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
