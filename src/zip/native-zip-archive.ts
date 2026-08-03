import type { ByteRange } from "../io/byte-range";
import { readAllBytes } from "../io/source";
import type { ZipArchive, ZipEntry } from "./zip-archive";

// A read-only ZIP reader over a ByteRange. It reads the central directory to
// locate entries, then reads and decompresses each entry on demand, so the
// whole archive need never be held. Field offsets are from the ZIP spec (PKWARE
// APPNOTE.TXT, 4.3).
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
const LOCAL_HEADER_SIZE = 30;
const MAX_COMMENT_SIZE = 0xffff;
const OVERFLOW_32 = 0xffff_ffff;

// A seekable source reads each entry in windows of this size rather than whole,
// so a large entry is never fully resident. In-memory reads ignore it, since a
// window is a view, not a copy.
const ENTRY_CHUNK_SIZE = 1 << 20;

interface EntryRecord {
  readonly method: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly localHeaderOffset: number;
}

export class NativeZipArchive implements ZipArchive {
  static async open(source: ByteRange): Promise<NativeZipArchive> {
    if (source.size < EOCD_MIN_SIZE) {
      throw new Error("Malformed zip: too short to contain a central directory");
    }
    return new NativeZipArchive(source, await readCentralDirectory(source));
  }

  private readonly entryList: readonly ZipEntry[];

  private constructor(
    private readonly source: ByteRange,
    private readonly records: ReadonlyMap<string, EntryRecord>,
  ) {
    this.entryList = [...records].map(([path, record]) => ({ path, size: record.uncompressedSize }));
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

    if (record.method === STORED) {
      return this.entryStream(record);
    }
    if (record.method === DEFLATE) {
      const stream = new DecompressionStream("deflate-raw");
      void this.inflate(stream.writable, record);
      return stream.readable;
    }
    throw new Error(`Unsupported compression method ${record.method} for ${path}`);
  }

  private entryStream(record: EntryRecord): ReadableStream<Uint8Array<ArrayBuffer>> {
    const chunks = readEntryChunks(this.source, record)[Symbol.asyncIterator]();
    return new ReadableStream<Uint8Array<ArrayBuffer>>({
      async pull(controller): Promise<void> {
        try {
          const { done, value } = await chunks.next();
          if (done) {
            controller.close();
          } else {
            controller.enqueue(value);
          }
        } catch (error) {
          controller.error(error);
        }
      },
      cancel(): void {
        void chunks.return?.(undefined);
      },
    });
  }

  private async inflate(writable: WritableStream<Uint8Array<ArrayBuffer>>, record: EntryRecord): Promise<void> {
    const writer = writable.getWriter();
    try {
      for await (const chunk of readEntryChunks(this.source, record)) {
        await writer.write(chunk);
      }
      await writer.close();
    } catch (error) {
      await writer.abort(error).catch(() => {});
    }
  }
}

async function readCentralDirectory(source: ByteRange): Promise<Map<string, EntryRecord>> {
  const tailSize = Math.min(source.size, EOCD_MIN_SIZE + MAX_COMMENT_SIZE);
  const tail = await source.read(source.size - tailSize, tailSize);
  const tailView = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);
  const eocd = findEndOfCentralDirectory(tailView, tail.length);

  const entryCount = tailView.getUint16(eocd + 10, true);
  const directorySize = tailView.getUint32(eocd + 12, true);
  const directoryOffset = tailView.getUint32(eocd + 16, true);
  if (entryCount === 0xffff || directorySize === OVERFLOW_32 || directoryOffset === OVERFLOW_32) {
    throw new Error("Zip64 archives are not supported");
  }

  const directory = await source.read(directoryOffset, directorySize);
  const view = new DataView(directory.buffer, directory.byteOffset, directory.byteLength);
  const decoder = new TextDecoder();
  const records = new Map<string, EntryRecord>();
  let pos = 0;
  for (let i = 0; i < entryCount; i++) {
    if (view.getUint32(pos, true) !== CENTRAL_HEADER_SIGNATURE) {
      throw new Error("Malformed zip: expected a central directory header");
    }

    const method = view.getUint16(pos + 10, true);
    const compressedSize = view.getUint32(pos + 20, true);
    const uncompressedSize = view.getUint32(pos + 24, true);
    const nameLength = view.getUint16(pos + 28, true);
    const extraLength = view.getUint16(pos + 30, true);
    const commentLength = view.getUint16(pos + 32, true);
    const localHeaderOffset = view.getUint32(pos + 42, true);

    if (compressedSize === OVERFLOW_32 || uncompressedSize === OVERFLOW_32 || localHeaderOffset === OVERFLOW_32) {
      throw new Error("Zip64 archives are not supported");
    }

    const name = decoder.decode(directory.subarray(pos + 46, pos + 46 + nameLength));
    records.set(name, { method, compressedSize, uncompressedSize, localHeaderOffset });

    pos += 46 + nameLength + extraLength + commentLength;
  }

  return records;
}

// The end of central directory record sits at the very end, unless a trailing
// comment pushes it earlier, so scan backwards for its signature.
function findEndOfCentralDirectory(view: DataView, length: number): number {
  for (let pos = length - EOCD_MIN_SIZE; pos >= 0; pos--) {
    if (view.getUint32(pos, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      return pos;
    }
  }
  throw new Error("Malformed zip: no end of central directory record");
}

async function* readEntryChunks(source: ByteRange, record: EntryRecord): AsyncIterable<Uint8Array<ArrayBuffer>> {
  const dataStart = await entryDataStart(source, record);
  for (let offset = 0; offset < record.compressedSize; offset += ENTRY_CHUNK_SIZE) {
    const length = Math.min(ENTRY_CHUNK_SIZE, record.compressedSize - offset);
    yield await source.read(dataStart + offset, length);
  }
}

// A local header repeats the name and extra fields with their own lengths,
// which can differ from the central directory, so the data offset is only known
// after reading them here.
async function entryDataStart(source: ByteRange, record: EntryRecord): Promise<number> {
  const header = await source.read(record.localHeaderOffset, LOCAL_HEADER_SIZE);
  const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
  if (view.getUint32(0, true) !== LOCAL_HEADER_SIGNATURE) {
    throw new Error("Malformed zip: expected a local file header");
  }
  const nameLength = view.getUint16(26, true);
  const extraLength = view.getUint16(28, true);
  return record.localHeaderOffset + LOCAL_HEADER_SIZE + nameLength + extraLength;
}
