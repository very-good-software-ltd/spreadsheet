export type CellValue =
  | { readonly type: "number"; readonly value: number }
  | { readonly type: "string"; readonly value: string }
  | { readonly type: "boolean"; readonly value: boolean }
  | { readonly type: "error"; readonly value: string };

export type Cell = { readonly ref: string } & CellValue;

export interface Row {
  readonly number: number;
  readonly cells: readonly Cell[];
}
