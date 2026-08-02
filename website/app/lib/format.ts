import type { Cell } from "very-good-spreadsheet";

export function columnLetter(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}

export function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(0)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function formatDate(value: Date): string {
  const iso = value.toISOString();
  return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : `${iso.slice(0, 10)} ${iso.slice(11, 16)}`;
}

export interface FormattedCell {
  text: string;
  align: "text-left" | "text-right";
  tone: string;
}

export function formatCell(cell: Cell): FormattedCell {
  switch (cell.type) {
    case "number":
      return { text: String(cell.value), align: "text-right", tone: "" };
    case "date":
      return { text: formatDate(cell.value), align: "text-right", tone: "text-emerald-700 dark:text-emerald-400" };
    case "boolean":
      return { text: cell.value ? "TRUE" : "FALSE", align: "text-left", tone: "text-amber-700 dark:text-amber-400" };
    case "error":
      return { text: cell.value, align: "text-left", tone: "text-red-600 dark:text-red-400" };
    default:
      return { text: cell.value, align: "text-left", tone: "" };
  }
}
