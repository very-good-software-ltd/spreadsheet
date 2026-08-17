import type { StoredZipEntry } from "./zip-archive";

/**
 * Assembles a ZIP container as a stream. Entries are declared first and produced
 * only as the output drains, so content never has to be held.
 */
export interface ZipWriter {
  /**
   * Place an entry from another archive unchanged, still compressed. Throws if
   * the path was already declared.
   */
  copy(entry: StoredZipEntry): void;

  /**
   * Declare an entry whose bytes are produced by `chunks` when the output
   * reaches it. Throws if the path was already declared.
   */
  add(path: string, chunks: () => AsyncIterable<Uint8Array>): void;

  /**
   * The archive as a stream, entries in declaration order. A failure in any
   * content source errors the stream, leaving a truncated archive that no reader
   * will accept. Throws if called more than once.
   */
  open(): ReadableStream<Uint8Array>;
}
