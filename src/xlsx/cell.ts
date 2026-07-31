export type CellValue =
  | { readonly type: "number"; readonly value: number }
  | { readonly type: "string"; readonly value: string };

export type Cell = { readonly ref: string } & CellValue;

export interface Row {
  readonly number: number;
  readonly cells: readonly Cell[];
}
