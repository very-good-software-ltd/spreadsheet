import { type BinarySource, readAllBytes } from "../io/source";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlReader } from "../xml/xml-reader";
import { openZip } from "../zip/open-zip";
import type { ZipArchive } from "../zip/zip-archive";
import { readSharedStrings } from "./read-shared-strings";
import { readWorksheetRefs, type WorksheetRef } from "./read-worksheet-refs";
import { Worksheet } from "./worksheet";

export class Workbook {
  static async open(source: BinarySource): Promise<Workbook> {
    const archive = openZip(await readAllBytes(source));
    const xml = createXmlReader();
    const refs = await readWorksheetRefs(archive, xml);
    const sharedStrings = await readSharedStrings(archive, xml);

    return new Workbook(archive, xml, refs, sharedStrings);
  }

  constructor(
    private readonly archive: ZipArchive,
    private readonly xml: XmlReader,
    private readonly refs: readonly WorksheetRef[],
    private readonly sharedStrings: readonly string[],
  ) {}

  get worksheetNames(): readonly string[] {
    return this.refs.map((ref) => ref.name);
  }

  worksheet(name: string): Worksheet {
    const ref = this.refs.find((candidate) => candidate.name === name);

    if (ref === undefined) {
      throw new Error(`Worksheet not found: ${name}`);
    }

    return new Worksheet(this.archive, this.xml, ref.path, this.sharedStrings);
  }
}
