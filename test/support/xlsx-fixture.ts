import { strToU8, zipSync } from "fflate";

type CellInput =
  | number
  | string
  | Date
  | { readonly boolean: boolean }
  | { readonly error: string }
  | { readonly inlineString: string }
  | { readonly formulaString: string }
  | { readonly rawType: string };

export interface SheetInput {
  readonly name: string;
  readonly rows: readonly (readonly CellInput[])[];
  readonly hidden?: boolean;
}

export interface WorkbookOptions {
  readonly date1904?: boolean;
}

const DAY_MS = 86_400_000;
const DATE_STYLE_INDEX = 1;

export function xlsx(sheets: readonly SheetInput[], options: WorkbookOptions = {}): Uint8Array {
  const date1904 = options.date1904 ?? false;

  const sharedStrings: string[] = [];
  const indexOfString = (text: string): number => {
    const existing = sharedStrings.indexOf(text);
    if (existing >= 0) {
      return existing;
    }
    sharedStrings.push(text);
    return sharedStrings.length - 1;
  };

  const cellXml = (ref: string, value: CellInput): string => {
    if (typeof value === "number") {
      return `<c r="${ref}"><v>${value}</v></c>`;
    }
    if (typeof value === "string") {
      return `<c r="${ref}" t="s"><v>${indexOfString(value)}</v></c>`;
    }
    if (value instanceof Date) {
      return `<c r="${ref}" s="${DATE_STYLE_INDEX}"><v>${dateToSerial(value, date1904)}</v></c>`;
    }
    if ("boolean" in value) {
      return `<c r="${ref}" t="b"><v>${value.boolean ? 1 : 0}</v></c>`;
    }
    if ("error" in value) {
      return `<c r="${ref}" t="e"><v>${value.error}</v></c>`;
    }
    if ("inlineString" in value) {
      return `<c r="${ref}" t="inlineStr"><is><t>${value.inlineString}</t></is></c>`;
    }
    if ("formulaString" in value) {
      return `<c r="${ref}" t="str"><v>${value.formulaString}</v></c>`;
    }
    return `<c r="${ref}" t="${value.rawType}"><v>0</v></c>`;
  };

  const files: Record<string, Uint8Array> = {};

  sheets.forEach((sheet, sheetIndex) => {
    const rowsXml = sheet.rows
      .map((cells, rowIndex) => {
        const cellsXml = cells
          .map((value, columnIndex) => cellXml(`${columnLetter(columnIndex)}${rowIndex + 1}`, value))
          .join("");
        return `<row r="${rowIndex + 1}">${cellsXml}</row>`;
      })
      .join("");

    files[`xl/worksheets/sheet${sheetIndex + 1}.xml`] = strToU8(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData></worksheet>`,
    );
  });

  const sheetElements = sheets
    .map((sheet, i) => {
      const state = sheet.hidden ? ' state="hidden"' : "";
      return `<sheet name="${sheet.name}" sheetId="${i + 1}" r:id="rId${i + 1}"${state}/>`;
    })
    .join("");
  const workbookProperties = date1904 ? '<workbookPr date1904="1"/>' : "";
  files["xl/workbook.xml"] = strToU8(
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${workbookProperties}<sheets>${sheetElements}</sheets></workbook>`,
  );

  const relElements = sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
    )
    .join("");
  files["xl/_rels/workbook.xml.rels"] = strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relElements}</Relationships>`,
  );

  const siElements = sharedStrings.map((text) => `<si><t>${text}</t></si>`).join("");
  files["xl/sharedStrings.xml"] = strToU8(
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${siElements}</sst>`,
  );

  // cellXfs index 1 uses built-in number format 14 (a date), which is what date cells reference.
  files["xl/styles.xml"] = strToU8(
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>`,
  );

  return zipSync(files);
}

function dateToSerial(date: Date, date1904: boolean): number {
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (date.getTime() - base) / DAY_MS;
}

function columnLetter(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}
