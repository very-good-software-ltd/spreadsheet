import { type BinarySource, readAllBytes } from "../io/source";
import { openZip } from "../zip/open-zip";
import type { ZipArchive } from "../zip/zip-archive";

export class Workbook {
  static async open(source: BinarySource): Promise<Workbook> {
    return new Workbook(openZip(await readAllBytes(source)));
  }

  constructor(readonly archive: ZipArchive) {}
}
