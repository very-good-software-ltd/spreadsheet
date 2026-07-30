import { type BinarySource, readAllBytes } from "../io/source";
import { createXmlReader } from "../xml/create-xml-reader";
import { openZip } from "../zip/open-zip";
import type { ZipArchive } from "../zip/zip-archive";
import { readWorksheetNames } from "./read-worksheet-names";

export class Workbook {
  static async open(source: BinarySource): Promise<Workbook> {
    const archive = openZip(await readAllBytes(source));
    const worksheetNames = await readWorksheetNames(archive, createXmlReader());

    return new Workbook(archive, worksheetNames);
  }

  constructor(
    readonly archive: ZipArchive,
    readonly worksheetNames: readonly string[],
  ) {}
}
