import type { Cell } from "./cell";

const CellTypeCode = {
  Number: "n",
  SharedString: "s",
  Boolean: "b",
  Error: "e",
  FormulaString: "str",
  InlineString: "inlineStr",
  Date: "d",
} as const;

export function interpretCell(ref: string, typeCode: string, text: string, sharedStrings: readonly string[]): Cell {
  // An absent t attribute means a number cell.
  const code = typeCode === "" ? CellTypeCode.Number : typeCode;

  switch (code) {
    case CellTypeCode.Number:
      return { ref, type: "number", value: Number(text) };
    case CellTypeCode.SharedString:
      return { ref, type: "string", value: sharedStrings[Number(text)] ?? "" };
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
