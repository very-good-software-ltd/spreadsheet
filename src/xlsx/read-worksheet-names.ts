import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";

const WORKBOOK_PART = "xl/workbook.xml";

export async function readWorksheetNames(archive: ZipArchive, xml: XmlReader): Promise<string[]> {
  const names: string[] = [];

  for await (const event of xml.read(archive.openStream(WORKBOOK_PART))) {
    if (event.type === "open" && event.name === "sheet") {
      const name = event.attributes["name"];
      if (name !== undefined) {
        names.push(name);
      }
    }
  }

  return names;
}
