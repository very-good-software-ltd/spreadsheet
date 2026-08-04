import type { XmlEvent, XmlReader } from "./xml/xml-reader";
import type { ZipArchive } from "./zip/zip-archive";

// Streams a part's XML events, tagging any read or parse failure with the part
// path so a malformed part reports which one, not a bare position from saxes.
export async function* readPart(archive: ZipArchive, xml: XmlReader, part: string): AsyncIterable<readonly XmlEvent[]> {
  try {
    yield* xml.read(archive.openStream(part));
  } catch (cause) {
    throw new Error(`Failed to read ${part}`, { cause });
  }
}
