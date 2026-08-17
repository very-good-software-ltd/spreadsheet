import { arrayBufferBacked } from "../io/byte-range";
import { crc32 } from "./crc32";
import { DOS_EPOCH_DATE, DOS_EPOCH_TIME } from "./dos-epoch";
import type { StoredZipEntry } from "./zip-archive";
import type { ZipWriter } from "./zip-writer";

// Field offsets and header layouts are from the ZIP spec (PKWARE APPNOTE.TXT,
// 4.3). Only what a spreadsheet needs is written: no Zip64, no encryption, and
// deflate for anything we produce ourselves.

const LOCAL_HEADER_SIGNATURE = 0x0403_4b50;
const CENTRAL_HEADER_SIGNATURE = 0x0201_4b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x0807_4b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x0605_4b50;

const LOCAL_HEADER_SIZE = 30;
const CENTRAL_HEADER_SIZE = 46;
const DATA_DESCRIPTOR_SIZE = 16;
const EOCD_SIZE = 22;

const DEFLATE = 8;
const VERSION_NEEDED_FOR_DEFLATE = 20;

const UTF8_PATH_FLAG = 0x0800;

// A zip header holds a size or an offset in 32 bits, and the entry count in 16.
// Past either, a file needs Zip64, which we do not write. Our reader throws on
// Zip64 rather than misread a file, so the writer throws rather than write one no
// reader will open. Handling files this large is its own piece of work.
const MAX_SIZE = 0xffff_ffff;
const MAX_ENTRIES = 0xffff;

// Bit 3 says the checksum and sizes are absent from the local header and follow
// the data instead. Deflating as we stream means we cannot know either until the
// data is written, and the local header comes first, so an entry we compress
// ourselves has to be described afterwards. The central directory, written last,
// always carries the real values, which is where a reader looks.
const DATA_DESCRIPTOR_FLAG = 0x0008;

interface DirectoryRecord {
  readonly path: Uint8Array;
  readonly method: number;
  readonly flags: number;
  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly dosTime: number;
  readonly dosDate: number;
  readonly localHeaderOffset: number;
}

type PendingEntry =
  | { readonly kind: "copy"; readonly path: string; readonly stored: StoredZipEntry }
  | { readonly kind: "add"; readonly path: string; readonly chunks: () => AsyncIterable<Uint8Array> };

export class NativeZipWriter implements ZipWriter {
  private readonly pending: PendingEntry[] = [];
  private readonly declared = new Set<string>();
  private opened = false;

  copy(entry: StoredZipEntry): void {
    // A copy declares its sizes up front, so this is knowable at the call.
    checkSize(entry.path, "compressed size", entry.compressedSize);
    checkSize(entry.path, "size", entry.uncompressedSize);
    this.declare(entry.path);
    this.pending.push({ kind: "copy", path: entry.path, stored: entry });
  }

  add(path: string, chunks: () => AsyncIterable<Uint8Array>): void {
    this.declare(path);
    this.pending.push({ kind: "add", path, chunks });
  }

  open(): ReadableStream<Uint8Array> {
    if (this.opened) {
      throw new Error("A zip writer can only be opened once");
    }
    this.opened = true;

    return toStream(writeArchive(this.pending));
  }

  private declare(path: string): void {
    if (this.declared.has(path)) {
      throw new Error(`Zip entry already declared: ${path}`);
    }
    if (this.declared.size >= MAX_ENTRIES) {
      throw new Error(
        `Cannot write ${MAX_ENTRIES + 1} entries: a zip lists at most ${MAX_ENTRIES} without Zip64, which we do not write`,
      );
    }
    this.declared.add(path);
  }
}

async function* writeArchive(entries: readonly PendingEntry[]): AsyncIterable<Uint8Array> {
  const encoder = new TextEncoder();
  const directory: DirectoryRecord[] = [];
  let offset = 0;

  for (const entry of entries) {
    const path = encoder.encode(entry.path);
    const localHeaderOffset = offset;
    checkSize(entry.path, "position in the archive", localHeaderOffset);

    if (entry.kind === "copy") {
      const { stored } = entry;
      // Only the UTF-8 marker is carried over. Everything else in the source's
      // flags describes how it was written, which is no longer true of this copy.
      const flags = stored.utf8Path ? UTF8_PATH_FLAG : 0;
      const record: DirectoryRecord = {
        path,
        method: stored.method,
        flags,
        crc32: stored.crc32,
        compressedSize: stored.compressedSize,
        uncompressedSize: stored.uncompressedSize,
        dosTime: stored.dosTime,
        dosDate: stored.dosDate,
        localHeaderOffset,
      };

      yield localHeader(record);
      offset += LOCAL_HEADER_SIZE + path.length;

      for await (const chunk of streamChunks(stored.bytes())) {
        offset += chunk.length;
        yield chunk;
      }

      directory.push(record);
      continue;
    }

    const flags = UTF8_PATH_FLAG | DATA_DESCRIPTOR_FLAG;
    yield localHeader({
      path,
      method: DEFLATE,
      flags,
      crc32: 0,
      compressedSize: 0,
      uncompressedSize: 0,
      dosTime: DOS_EPOCH_TIME,
      dosDate: DOS_EPOCH_DATE,
      localHeaderOffset,
    });
    offset += LOCAL_HEADER_SIZE + path.length;

    const source = { crc32: 0, size: 0 };
    let compressedSize = 0;

    for await (const chunk of deflate(entry.chunks(), source)) {
      compressedSize += chunk.length;
      offset += chunk.length;
      yield chunk;
    }

    // An entry we compress ourselves only reveals its sizes once it is written,
    // so unlike a copy this cannot be caught at the call. The stream errors and
    // leaves a truncated archive, which is the same failure as any other late one.
    checkSize(entry.path, "compressed size", compressedSize);
    checkSize(entry.path, "size", source.size);

    yield dataDescriptor(source.crc32, compressedSize, source.size);
    offset += DATA_DESCRIPTOR_SIZE;

    directory.push({
      path,
      method: DEFLATE,
      flags,
      crc32: source.crc32,
      compressedSize,
      uncompressedSize: source.size,
      dosTime: DOS_EPOCH_TIME,
      dosDate: DOS_EPOCH_DATE,
      localHeaderOffset,
    });
  }

  const directoryOffset = offset;
  checkSize("the central directory", "position in the archive", directoryOffset);
  let directorySize = 0;

  for (const record of directory) {
    const header = centralHeader(record);
    directorySize += header.length;
    yield header;
  }

  yield endOfCentralDirectory(directory.length, directorySize, directoryOffset);
}

// Deflates `chunks`, accumulating the checksum and byte count of the input into
// `source` as it goes, so both are final by the time the output is drained.
async function* deflate(
  chunks: AsyncIterable<Uint8Array>,
  source: { crc32: number; size: number },
): AsyncIterable<Uint8Array> {
  const compression = new CompressionStream("deflate-raw");
  const failure = { error: undefined as unknown };

  void (async (): Promise<void> => {
    const writer = compression.writable.getWriter();
    try {
      for await (const chunk of chunks) {
        source.crc32 = crc32(chunk, source.crc32);
        source.size += chunk.length;
        await writer.write(arrayBufferBacked(chunk));
      }
      await writer.close();
    } catch (error) {
      failure.error = error;
      await writer.abort(error).catch(() => {});
    }
  })();

  yield* streamChunks(compression.readable);

  if (failure.error !== undefined) {
    throw failure.error;
  }
}

async function* streamChunks(stream: ReadableStream<Uint8Array>): AsyncIterable<Uint8Array> {
  const reader = stream.getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        return;
      }
      yield value;
    }
  } finally {
    reader.releaseLock();
  }
}

function toStream(chunks: AsyncIterable<Uint8Array>): ReadableStream<Uint8Array> {
  const iterator = chunks[Symbol.asyncIterator]();

  return new ReadableStream<Uint8Array>({
    async pull(controller): Promise<void> {
      try {
        const { done, value } = await iterator.next();
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
      void iterator.return?.(undefined);
    },
  });
}

function checkSize(what: string, field: string, value: number): void {
  if (value > MAX_SIZE) {
    throw new Error(
      `Cannot write ${what}: its ${field} is ${value} bytes, past the ${MAX_SIZE + 1} a zip addresses without Zip64, which we do not write`,
    );
  }
}

function localHeader(record: DirectoryRecord): Uint8Array {
  const header = new Uint8Array(LOCAL_HEADER_SIZE + record.path.length);
  const view = new DataView(header.buffer);

  view.setUint32(0, LOCAL_HEADER_SIGNATURE, true);
  view.setUint16(4, VERSION_NEEDED_FOR_DEFLATE, true);
  view.setUint16(6, record.flags, true);
  view.setUint16(8, record.method, true);
  view.setUint16(10, record.dosTime, true);
  view.setUint16(12, record.dosDate, true);
  view.setUint32(14, record.crc32, true);
  view.setUint32(18, record.compressedSize, true);
  view.setUint32(22, record.uncompressedSize, true);
  view.setUint16(26, record.path.length, true);
  view.setUint16(28, 0, true);
  header.set(record.path, LOCAL_HEADER_SIZE);

  return header;
}

function dataDescriptor(entryCrc32: number, compressedSize: number, uncompressedSize: number): Uint8Array {
  const descriptor = new Uint8Array(DATA_DESCRIPTOR_SIZE);
  const view = new DataView(descriptor.buffer);

  view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
  view.setUint32(4, entryCrc32, true);
  view.setUint32(8, compressedSize, true);
  view.setUint32(12, uncompressedSize, true);

  return descriptor;
}

function centralHeader(record: DirectoryRecord): Uint8Array {
  const header = new Uint8Array(CENTRAL_HEADER_SIZE + record.path.length);
  const view = new DataView(header.buffer);

  view.setUint32(0, CENTRAL_HEADER_SIGNATURE, true);
  view.setUint16(4, VERSION_NEEDED_FOR_DEFLATE, true);
  view.setUint16(6, VERSION_NEEDED_FOR_DEFLATE, true);
  view.setUint16(8, record.flags, true);
  view.setUint16(10, record.method, true);
  view.setUint16(12, record.dosTime, true);
  view.setUint16(14, record.dosDate, true);
  view.setUint32(16, record.crc32, true);
  view.setUint32(20, record.compressedSize, true);
  view.setUint32(24, record.uncompressedSize, true);
  view.setUint16(28, record.path.length, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, record.localHeaderOffset, true);
  header.set(record.path, CENTRAL_HEADER_SIZE);

  return header;
}

function endOfCentralDirectory(entryCount: number, directorySize: number, directoryOffset: number): Uint8Array {
  const eocd = new Uint8Array(EOCD_SIZE);
  const view = new DataView(eocd.buffer);

  view.setUint32(0, END_OF_CENTRAL_DIRECTORY_SIGNATURE, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, entryCount, true);
  view.setUint16(10, entryCount, true);
  view.setUint32(12, directorySize, true);
  view.setUint32(16, directoryOffset, true);
  view.setUint16(20, 0, true);

  return eocd;
}
