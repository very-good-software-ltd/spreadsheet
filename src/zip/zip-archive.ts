/**
 * A stream-based read-only view over the entries of a ZIP container.
 */
export interface ZipArchive {
  /** All entries found in the central directory. */
  entries(): readonly ZipEntry[];

  /** Whether an entry with this exact path exists. */
  has(path: string): boolean;

  /** Read an entry's full decompressed bytes. Throws if the path is missing. */
  read(path: string): Promise<Uint8Array>;

  /** Open an entry's decompressed bytes as a stream. Throws if the path is missing. */
  openStream(path: string): ReadableStream<Uint8Array>;

  /**
   * An entry as it is stored, still compressed, with the header fields needed to
   * place it in another archive unchanged. Throws if the path is missing.
   */
  storedEntry(path: string): StoredZipEntry;
}

/**
 * An entry's bytes exactly as they sit in the archive, alongside the fields a
 * writer must repeat to keep them readable. Copying one needs no decompression
 * and no checksumming, so a part nobody understands survives intact.
 */
export interface StoredZipEntry {
  readonly path: string;

  /** The zip compression method code, to be written through unchanged. */
  readonly method: number;

  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;

  /** The MS-DOS packed modification time and date fields, preserved verbatim. */
  readonly dosTime: number;
  readonly dosDate: number;

  /** Whether the source flagged the entry path as UTF-8. */
  readonly utf8Path: boolean;

  /** The still-compressed bytes. */
  bytes(): ReadableStream<Uint8Array>;
}

export interface ZipEntry {
  /** The entry path within the archive, for example `xl/worksheets/sheet1.xml`. */
  readonly path: string;

  /** Uncompressed size in bytes. */
  readonly size: number;
}
