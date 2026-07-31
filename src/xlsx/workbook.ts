import { type BinarySource, readAllBytes } from "../io/source";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlReader } from "../xml/xml-reader";
import { openZip } from "../zip/open-zip";
import type { ZipArchive } from "../zip/zip-archive";
import type { CellContext } from "./interpret-cell";
import { readSharedStrings } from "./read-shared-strings";
import { readStyles } from "./read-styles";
import { readWorkbook, type WorksheetRef } from "./read-workbook";
import { Worksheet } from "./worksheet";

export interface WorksheetInfo {
  readonly name: string;
  readonly hidden: boolean;
}

export class Workbook {
  static async open(source: BinarySource): Promise<Workbook> {
    const archive = openZip(await readAllBytes(source));
    const xml = createXmlReader();
    const { worksheets, date1904 } = await readWorkbook(archive, xml);
    const sharedStrings = await readSharedStrings(archive, xml);
    const styles = await readStyles(archive, xml);

    return new Workbook(archive, xml, worksheets, { sharedStrings, styles, date1904 });
  }

  constructor(
    private readonly archive: ZipArchive,
    private readonly xml: XmlReader,
    private readonly refs: readonly WorksheetRef[],
    private readonly context: CellContext,
  ) {}

  get worksheets(): readonly WorksheetInfo[] {
    return this.refs.map((ref) => ({ name: ref.name, hidden: ref.hidden }));
  }

  get worksheetNames(): readonly string[] {
    return this.refs.map((ref) => ref.name);
  }

  worksheet(name: string): Worksheet {
    const ref = this.refs.find((candidate) => candidate.name === name);

    if (ref === undefined) {
      throw new Error(`Worksheet not found: ${name}`);
    }

    return new Worksheet(this.archive, this.xml, ref.path, this.context);
  }
}
