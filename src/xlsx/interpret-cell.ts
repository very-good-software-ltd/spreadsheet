import type { Cell } from "./cell";
import { serialToDate } from "./date";
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

export function interpretCell(
  ref: string,
  typeCode: string,
  styleIndex: number | undefined,
  text: string,
  context: CellContext,
): Cell {
  // An absent t attribute means a number cell.
  const code = typeCode === "" ? CellTypeCode.Number : typeCode;

  switch (code) {
    case CellTypeCode.Number:
      if (styleIndex !== undefined && context.styles.isDateStyle(styleIndex)) {
        return { ref, type: "date", value: serialToDate(Number(text), context.date1904) };
      }
      return { ref, type: "number", value: Number(text) };
    case CellTypeCode.SharedString:
      return { ref, type: "string", value: context.sharedStrings[Number(text)] ?? "" };
    case CellTypeCode.FormulaString:
    case CellTypeCode.InlineString:
      return { ref, type: "string", value: text };
    case CellTypeCode.Boolean:
      return { ref, type: "boolean", value: text !== "0" };
    case CellTypeCode.Error:
      return { ref, type: "error", value: text };
    default:
      throw new Error(`Unsupported cell type "${typeCode}" at ${ref}`);
  }
}
