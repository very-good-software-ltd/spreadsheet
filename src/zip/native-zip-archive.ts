import { readAllBytes } from "../io/source";
import type { ZipArchive, ZipEntry } from "./zip-archive";

// A read-only ZIP reader over the whole archive bytes. It parses the central
// directory to locate entries, then decompresses each entry on demand through
// the platform's DecompressionStream, so a large entry is never fully held in
// memory. Field offsets are from the ZIP spec (PKWARE APPNOTE.TXT, 4.3).
//
// Not supported, each throws rather than misreading: Zip64 (sizes or offsets
// that overflow 32 bits), encryption, and compression methods other than
// stored (0) and deflate (8). xlsx from real tools stays within this.

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x0605_4b50;
const CENTRAL_HEADER_SIGNATURE = 0x0201_4b50;
const LOCAL_HEADER_SIGNATURE = 0x0403_4b50;

const STORED = 0;
const DEFLATE = 8;

const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_SIZE = 0xffff;
const OVERFLOW_32 = 0xffff_ffff;

interface EntryRecord {
  readonly method: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly localHeaderOffset: number;
}

export class NativeZipArchive implements ZipArchive {
  private readonly view: DataView;
  private readonly records: ReadonlyMap<string, EntryRecord>;
  private readonly entryList: readonly ZipEntry[];

  constructor(private readonly bytes: Uint8Array) {
    if (bytes.length < EOCD_MIN_SIZE) {
      throw new Error("Malformed zip: too short to contain a central directory");
    }
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.records = this.readCentralDirectory();
    this.entryList = [...this.records].map(([path, record]) => ({
      path,
      size: record.uncompressedSize,
    }));
  }

  entries(): readonly ZipEntry[] {
    return this.entryList;
  }

  has(path: string): boolean {
    return this.records.has(path);
  }

  read(path: string): Promise<Uint8Array> {
    return readAllBytes(this.openStream(path));
  }

  openStream(path: string): ReadableStream<Uint8Array> {
    const record = this.records.get(path);
    if (record === undefined) {
      throw new Error(`Zip entry not found: ${path}`);
    }

    const compressed = this.compressedData(record);

    if (record.method === STORED) {
      return new ReadableStream<Uint8Array>({
        start(controller): void {
          controller.enqueue(compressed);
          controller.close();
        },
      });
    }
    if (record.method === DEFLATE) {
      const stream = new DecompressionStream("deflate-raw");
      const writer = stream.writable.getWriter();
      // DecompressionStream needs an ArrayBuffer-backed view, so copy the slice
      // off the source buffer, which the type system treats as possibly shared.
      // The write can only fail as the consumer pulls, where the readable side
      // reports it, so this catch just stops that error from also going unhandled.
      void writer
        .write(new Uint8Array(compressed))
        .then(() => writer.close())
        .catch(() => {});
      return stream.readable;
    }
    throw new Error(`Unsupported compression method ${record.method} for ${path}`);
  }

  private readCentralDirectory(): Map<string, EntryRecord> {
    const eocd = this.findEndOfCentralDirectory();
    const entryCount = this.view.getUint16(eocd + 10, true);
    if (entryCount === 0xffff) {
      throw new Error("Zip64 archives are not supported");
    }

    const records = new Map<string, EntryRecord>();
    let pos = this.view.getUint32(eocd + 16, true);
    for (let i = 0; i < entryCount; i++) {
      if (this.view.getUint32(pos, true) !== CENTRAL_HEADER_SIGNATURE) {
        throw new Error("Malformed zip: expected a central directory header");
      }

      const method = this.view.getUint16(pos + 10, true);
      const compressedSize = this.view.getUint32(pos + 20, true);
      const uncompressedSize = this.view.getUint32(pos + 24, true);
      const nameLength = this.view.getUint16(pos + 28, true);
      const extraLength = this.view.getUint16(pos + 30, true);
      const commentLength = this.view.getUint16(pos + 32, true);
      const localHeaderOffset = this.view.getUint32(pos + 42, true);

      if (compressedSize === OVERFLOW_32 || uncompressedSize === OVERFLOW_32 || localHeaderOffset === OVERFLOW_32) {
        throw new Error("Zip64 archives are not supported");
      }

      const name = this.decodeName(pos + 46, nameLength);
      records.set(name, { method, compressedSize, uncompressedSize, localHeaderOffset });

      pos += 46 + nameLength + extraLength + commentLength;
    }

    return records;
  }

  // The end of central directory record sits at the very end, unless a trailing
  // comment pushes it earlier, so scan backwards for its signature.
  private findEndOfCentralDirectory(): number {
    const earliest = Math.max(0, this.bytes.length - EOCD_MIN_SIZE - MAX_COMMENT_SIZE);
    for (let pos = this.bytes.length - EOCD_MIN_SIZE; pos >= earliest; pos--) {
      if (this.view.getUint32(pos, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
        return pos;
      }
    }
    throw new Error("Malformed zip: no end of central directory record");
  }

  // A local header repeats the name and extra fields with their own lengths,
  // which can differ from the central directory, so the data offset is only
  // known after reading them here.
  private compressedData(record: EntryRecord): Uint8Array {
    const offset = record.localHeaderOffset;
    if (this.view.getUint32(offset, true) !== LOCAL_HEADER_SIGNATURE) {
      throw new Error("Malformed zip: expected a local file header");
    }
    const nameLength = this.view.getUint16(offset + 26, true);
    const extraLength = this.view.getUint16(offset + 28, true);
    const dataStart = offset + 30 + nameLength + extraLength;
    return this.bytes.subarray(dataStart, dataStart + record.compressedSize);
  }

  private decodeName(start: number, length: number): string {
    return new TextDecoder().decode(this.bytes.subarray(start, start + length));
  }
}
