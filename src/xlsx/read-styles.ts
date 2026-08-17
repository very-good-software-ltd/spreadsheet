import { readPart } from "../read-part";
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

// Ids below this are reserved for the formats every spreadsheet has built in, so
// a format a file defines itself takes an id from here up.
const FIRST_CUSTOM_FORMAT_ID = 164;

// Built-in number formats that render a number as a date or time.
const BUILTIN_DATE_FORMAT_IDS: ReadonlySet<number> = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

export class Styles {
  constructor(
    private readonly cellFormats: readonly Readonly<Record<string, string>>[],
    private readonly customFormatCodes: ReadonlyMap<number, string>,
    /** Whether the part declares a number format list to append to. */
    readonly hasNumberFormats: boolean,
  ) {}

  /** How many cell formats the part declares, which is the next free index. */
  get cellFormatCount(): number {
    return this.cellFormats.length;
  }

  /**
   * The attributes of a cell format, so a caller can derive a new one from it
   * and keep its font, fill and border. Absent if the index is not in the part.
   */
  cellFormat(styleIndex: number): Readonly<Record<string, string>> | undefined {
    return this.cellFormats[styleIndex];
  }

  /**
   * A number format id no format in the part uses. Ids below 164 are reserved for
   * the built-in formats, so a new one starts there.
   */
  nextCustomFormatId(): number {
    return Math.max(FIRST_CUSTOM_FORMAT_ID, ...[...this.customFormatCodes.keys()].map((id) => id + 1));
  }

  isDateStyle(styleIndex: number): boolean {
    const format = this.cellFormats[styleIndex];

    if (format === undefined) {
      return false;
    }

    const numberFormatId = Number(format[Attribute.NumberFormatId] ?? "0");

    if (BUILTIN_DATE_FORMAT_IDS.has(numberFormatId)) {
      return true;
    }

    const code = this.customFormatCodes.get(numberFormatId);

    return code !== undefined && looksLikeDate(code);
  }
}

export async function readStyles(archive: ZipArchive, xml: XmlReader): Promise<Styles> {
  const cellFormats: Readonly<Record<string, string>>[] = [];
  const customFormatCodes = new Map<number, string>();

  if (!archive.has(STYLES_PART)) {
    return new Styles(cellFormats, customFormatCodes, false);
  }

  let inCellFormats = false;
  let inNumberFormats = false;
  let hasNumberFormats = false;

  for await (const batch of readPart(archive, xml, STYLES_PART)) {
    for (const event of batch) {
      if (event.type === "open") {
        if (event.name === Element.CellFormats) {
          inCellFormats = true;
        } else if (event.name === Element.NumberFormats) {
          inNumberFormats = true;
          hasNumberFormats = true;
        } else if (event.name === Element.Format && inCellFormats) {
          cellFormats.push(event.attributes);
        } else if (event.name === Element.NumberFormat && inNumberFormats) {
          const id = event.attributes[Attribute.NumberFormatId];
          const code = event.attributes[Attribute.FormatCode];

          if (id !== undefined && code !== undefined) {
            customFormatCodes.set(Number(id), code);
          }
        }
      } else if (event.type === "close" && event.name === Element.CellFormats) {
        inCellFormats = false;
      } else if (event.type === "close" && event.name === Element.NumberFormats) {
        inNumberFormats = false;
      }
    }
  }

  return new Styles(cellFormats, customFormatCodes, hasNumberFormats);
}

function looksLikeDate(formatCode: string): boolean {
  const withoutLiterals = formatCode
    .replace(/\\./g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, "");
  return /[dhmsy]/i.test(withoutLiterals);
}
