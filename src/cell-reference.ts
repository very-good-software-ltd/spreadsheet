const UPPERCASE_A = 65;
const UPPERCASE_Z = 90;
const LETTERS_IN_ALPHABET = 26;

/**
 * The zero-based column index of a reference like `"C3"`, so `"A1"` is 0 and
 * `"AA1"` is 26.
 */
export function columnIndexOf(ref: string): number {
  let index = 0;

  for (const char of ref) {
    const code = char.charCodeAt(0);
    if (code < UPPERCASE_A || code > UPPERCASE_Z) {
      break;
    }
    index = index * LETTERS_IN_ALPHABET + (code - UPPERCASE_A + 1);
  }

  return index - 1;
}

/** The one-based row number of a reference like `"C3"`. */
export function rowNumberOf(ref: string): number {
  return Number(ref.replace(/^[A-Z]+/, ""));
}

/**
 * The column letters for a zero-based index, so 0 is `"A"` and 26 is `"AA"`.
 *
 * The letters are not plain base-26. There is no zero digit, so `"Z"` is followed
 * by `"AA"` rather than by a two-letter form of 26, which is why each step down
 * subtracts one before dividing.
 */
export function columnLetters(columnIndex: number): string {
  let remaining = columnIndex;
  let letters = "";

  do {
    letters = String.fromCharCode(UPPERCASE_A + (remaining % LETTERS_IN_ALPHABET)) + letters;
    remaining = Math.floor(remaining / LETTERS_IN_ALPHABET) - 1;
  } while (remaining >= 0);

  return letters;
}

/** A reference like `"C3"` from a one-based row and a zero-based column. */
export function cellReference(row: number, columnIndex: number): string {
  return `${columnLetters(columnIndex)}${row}`;
}
