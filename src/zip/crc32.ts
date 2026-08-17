// CRC-32/ISO-HDLC, the checksum every zip entry header carries. The reversed
// polynomial 0xEDB88320 and the inverted start and end state are the zip
// spec's, not a choice (PKWARE APPNOTE.TXT, 4.4.7).
const POLYNOMIAL = 0xedb8_8320;

const TABLE = buildTable();

function buildTable(): Uint32Array {
  const table = new Uint32Array(256);

  for (let byte = 0; byte < 256; byte++) {
    let value = byte;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? (value >>> 1) ^ POLYNOMIAL : value >>> 1;
    }
    table[byte] = value;
  }

  return table;
}

/**
 * The CRC-32 of `bytes`, continuing from `previous` so a stream can be
 * checksummed chunk by chunk. Pass the result of the last call as `previous`,
 * and omit it for the first chunk.
 */
export function crc32(bytes: Uint8Array, previous = 0): number {
  let state = ~previous;

  for (const byte of bytes) {
    // The table is 256 wide and the index is masked to a byte, so this cannot
    // miss, but the index signature does not know that.
    state = (state >>> 8) ^ (TABLE[(state ^ byte) & 0xff] as number);
  }

  return ~state >>> 0;
}
