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
}

export interface ZipEntry {
  /** The entry path within the archive, for example `xl/worksheets/sheet1.xml`. */
  readonly path: string;

  /** Uncompressed size in bytes. */
  readonly size: number;
}
