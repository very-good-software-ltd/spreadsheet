export type CellValue =
  | { readonly type: "number"; readonly value: number }
  | { readonly type: "string"; readonly value: string }
  | { readonly type: "boolean"; readonly value: boolean }
  | { readonly type: "error"; readonly value: string }
  | { readonly type: "date"; readonly value: Date };

export type Cell = { readonly ref: string; readonly columnIndex: number } & CellValue;
