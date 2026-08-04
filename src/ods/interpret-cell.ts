import type { CellValue, ResolvedValue } from "../cell";
import { parseIsoDate } from "../iso-date";

const Attribute = {
  ValueType: "office:value-type",
  Value: "office:value",
  DateValue: "office:date-value",
  BooleanValue: "office:boolean-value",
  TimeValue: "office:time-value",
  Formula: "table:formula",
} as const;

const ValueType = {
  Float: "float",
  Percentage: "percentage",
  Currency: "currency",
  String: "string",
  Boolean: "boolean",
  Date: "date",
  Time: "time",
} as const;

export interface OdsCell {
  readonly valueType: string;
  readonly formula: string | null;
  readonly numberValue: string;
  readonly dateValue: string;
  readonly booleanValue: string;
  readonly timeValue: string;
  text: string;
}

export function emptyCell(): OdsCell {
  return cellFrom({}, false);
}

export function cellFrom(attributes: Record<string, string>, covered: boolean): OdsCell {
  // A covered cell is merged away, so it holds no value of its own.
  return {
    valueType: covered ? "" : (attributes[Attribute.ValueType] ?? ""),
    formula: covered ? null : (attributes[Attribute.Formula] ?? null),
    numberValue: attributes[Attribute.Value] ?? "",
    dateValue: attributes[Attribute.DateValue] ?? "",
    booleanValue: attributes[Attribute.BooleanValue] ?? "",
    timeValue: attributes[Attribute.TimeValue] ?? "",
    text: "",
  };
}

export function isEmpty(cell: OdsCell): boolean {
  return cell.valueType === "" && cell.formula === null;
}

export function interpretOdsCell(cell: OdsCell): CellValue {
  const resolved = resolvedValue(cell);
  if (cell.formula !== null) {
    return { type: "formula", value: cell.formula, cachedValue: resolved };
  }
  if (resolved === null) {
    throw new Error("ods cell has neither a value nor a formula");
  }
  return resolved;
}

function resolvedValue(cell: OdsCell): ResolvedValue | null {
  switch (cell.valueType) {
    case "":
      return null;
    case ValueType.Float:
    case ValueType.Percentage:
    case ValueType.Currency:
      return { type: "number", value: Number(cell.numberValue) };
    case ValueType.String:
      return { type: "string", value: cell.text };
    case ValueType.Boolean:
      return { type: "boolean", value: cell.booleanValue === "true" };
    case ValueType.Date:
      return { type: "date", value: parseIsoDate(cell.dateValue) };
    // ODF times are ISO 8601 durations with no place in the value union, so the
    // raw duration is surfaced as a string.
    case ValueType.Time:
      return { type: "string", value: cell.timeValue };
    default:
      throw new Error(`Unsupported ods cell value type "${cell.valueType}"`);
  }
}
