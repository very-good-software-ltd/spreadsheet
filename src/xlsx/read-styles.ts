import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";

const STYLES_PART = "xl/styles.xml";

const Element = {
  NumberFormats: "numFmts",
  NumberFormat: "numFmt",
  CellFormats: "cellXfs",
  Format: "xf",
} as const;

const Attribute = {
  NumberFormatId: "numFmtId",
  FormatCode: "formatCode",
} as const;

// Built-in number formats that render a number as a date or time.
const BUILTIN_DATE_FORMAT_IDS: ReadonlySet<number> = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

export class Styles {
  constructor(
    private readonly cellNumberFormatIds: readonly number[],
    private readonly customFormatCodes: ReadonlyMap<number, string>,
  ) {}

  isDateStyle(styleIndex: number): boolean {
    const numberFormatId = this.cellNumberFormatIds[styleIndex];

    if (numberFormatId === undefined) {
      return false;
    }

    if (BUILTIN_DATE_FORMAT_IDS.has(numberFormatId)) {
      return true;
    }

    const code = this.customFormatCodes.get(numberFormatId);

    return code !== undefined && looksLikeDate(code);
  }
}

export async function readStyles(archive: ZipArchive, xml: XmlReader): Promise<Styles> {
  const cellNumberFormatIds: number[] = [];
  const customFormatCodes = new Map<number, string>();

  if (!archive.has(STYLES_PART)) {
    return new Styles(cellNumberFormatIds, customFormatCodes);
  }

  let inCellFormats = false;

  for await (const event of xml.read(archive.openStream(STYLES_PART))) {
    if (event.type === "open") {
      if (event.name === Element.CellFormats) {
        inCellFormats = true;
      } else if (event.name === Element.Format && inCellFormats) {
        cellNumberFormatIds.push(Number(event.attributes[Attribute.NumberFormatId] ?? "0"));
      } else if (event.name === Element.NumberFormat) {
        const id = event.attributes[Attribute.NumberFormatId];
        const code = event.attributes[Attribute.FormatCode];

        if (id !== undefined && code !== undefined) {
          customFormatCodes.set(Number(id), code);
        }
      }
    } else if (event.type === "close" && event.name === Element.CellFormats) {
      inCellFormats = false;
    }
  }

  return new Styles(cellNumberFormatIds, customFormatCodes);
}

function looksLikeDate(formatCode: string): boolean {
  const withoutLiterals = formatCode
    .replace(/\\./g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, "");
  return /[dhmsy]/i.test(withoutLiterals);
}
