import { crc32 } from "./crc32";
import { DOS_EPOCH_DATE, DOS_EPOCH_TIME } from "./dos-epoch";
import type { StoredZipEntry, ZipArchive, ZipEntry } from "./zip-archive";

// Entries are held uncompressed, so a reader hands the bytes straight back and a
// writer copying one through has nothing to inflate or deflate.
const STORED = 0;

/** A ZIP archive assembled in memory rather than read from bytes. */
export class MemoryZipArchive implements ZipArchive {
  static ofText(parts: Readonly<Record<string, string>>): MemoryZipArchive {
    const encoder = new TextEncoder();
    return new MemoryZipArchive(new Map(Object.entries(parts).map(([path, text]) => [path, encoder.encode(text)])));
  }

  constructor(private readonly parts: ReadonlyMap<string, Uint8Array>) {}

  entries(): readonly ZipEntry[] {
    return [...this.parts].map(([path, bytes]) => ({ path, size: bytes.length }));
  }

  has(path: string): boolean {
    return this.parts.has(path);
  }

  read(path: string): Promise<Uint8Array> {
    return Promise.resolve(this.bytesOf(path));
  }

  openStream(path: string): ReadableStream<Uint8Array> {
    const bytes = this.bytesOf(path);

    return new ReadableStream<Uint8Array>({
      start(controller): void {
        controller.enqueue(bytes);
        controller.close();
      },
    });
  }

  storedEntry(path: string): StoredZipEntry {
    const bytes = this.bytesOf(path);

    return {
      path,
      method: STORED,
      crc32: crc32(bytes),
      compressedSize: bytes.length,
      uncompressedSize: bytes.length,
      dosTime: DOS_EPOCH_TIME,
      dosDate: DOS_EPOCH_DATE,
      utf8Path: true,
      bytes: () => this.openStream(path),
    };
  }

  private bytesOf(path: string): Uint8Array {
    const bytes = this.parts.get(path);
    if (bytes === undefined) {
      throw new Error(`Zip entry not found: ${path}`);
    }
    return bytes;
  }
}
