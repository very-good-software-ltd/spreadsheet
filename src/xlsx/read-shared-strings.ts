import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { readPart } from "./read-part";

const SHARED_STRINGS_PART = "xl/sharedStrings.xml";

const Element = {
  StringItem: "si",
  Text: "t",
} as const;

export async function readSharedStrings(archive: ZipArchive, xml: XmlReader): Promise<string[]> {
  const strings: string[] = [];
  if (!archive.has(SHARED_STRINGS_PART)) {
    return strings;
  }

  let current: string | null = null;
  let inText = false;

  for await (const event of readPart(archive, xml, SHARED_STRINGS_PART)) {
    switch (event.type) {
      case "open":
        if (event.name === Element.StringItem) {
          current = "";
        } else if (event.name === Element.Text) {
          inText = true;
        }
        break;
      case "text":
        if (inText && current !== null) {
          current += event.text;
        }
        break;
      case "close":
        if (event.name === Element.Text) {
          inText = false;
        } else if (event.name === Element.StringItem) {
          strings.push(current ?? "");
          current = null;
        }
        break;
    }
  }

  return strings;
}
