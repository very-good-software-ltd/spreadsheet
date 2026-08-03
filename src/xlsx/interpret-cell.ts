import type { CellValue } from "../cell";
import { parseIsoDate, serialToDate } from "./date";
import type { Styles } from "./read-styles";

const CellTypeCode = {
  Number: "n",
  SharedString: "s",
  Boolean: "b",
  Error: "e",
  FormulaString: "str",
  InlineString: "inlineStr",
  Date: "d",
} as const;

export interface CellContext {
  readonly sharedStrings: readonly string[];
  readonly styles: Styles;
  readonly date1904: boolean;
}

export function interpretCellValue(
  ref: string,
  typeCode: string,
  styleIndex: number | undefined,
  text: string,
  context: CellContext,
): CellValue {
  // An absent t attribute means a number cell.
  const code = typeCode === "" ? CellTypeCode.Number : typeCode;

  switch (code) {
    case CellTypeCode.Number:
      if (styleIndex !== undefined && context.styles.isDateStyle(styleIndex)) {
        return { type: "date", value: serialToDate(Number(text), context.date1904) };
      }
      return { type: "number", value: Number(text) };
    case CellTypeCode.SharedString:
      return { type: "string", value: context.sharedStrings[Number(text)] ?? "" };
    case CellTypeCode.FormulaString:
    case CellTypeCode.InlineString:
      return { type: "string", value: text };
    case CellTypeCode.Boolean:
      return { type: "boolean", value: text !== "0" };
    case CellTypeCode.Error:
      return { type: "error", value: text };
    case CellTypeCode.Date:
      return { type: "date", value: parseIsoDate(text) };
    default:
      throw new Error(`Unsupported cell type "${typeCode}" at ${ref}`);
  }
}
