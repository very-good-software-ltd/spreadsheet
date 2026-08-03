import { type BinarySource, readAllBytes } from "./source";

export interface ByteRange {
  readonly size: number;
  read(offset: number, length: number): Promise<Uint8Array<ArrayBuffer>>;
}

export class BytesByteRange implements ByteRange {
  private readonly bytes: Uint8Array<ArrayBuffer>;

  constructor(bytes: Uint8Array) {
    this.bytes = arrayBufferBacked(bytes);
  }

  get size(): number {
    return this.bytes.length;
  }

  read(offset: number, length: number): Promise<Uint8Array<ArrayBuffer>> {
    return Promise.resolve(this.bytes.subarray(offset, offset + length));
  }
}

export class BlobByteRange implements ByteRange {
  constructor(private readonly blob: Blob) {}

  get size(): number {
    return this.blob.size;
  }

  async read(offset: number, length: number): Promise<Uint8Array<ArrayBuffer>> {
    return new Uint8Array(await this.blob.slice(offset, offset + length).arrayBuffer());
  }
}

export async function toByteRange(source: BinarySource | Blob): Promise<ByteRange> {
  if (source instanceof Blob) {
    return new BlobByteRange(source);
  }
  return new BytesByteRange(await readAllBytes(source));
}

// DecompressionStream's writer only accepts an ArrayBuffer-backed view, so the
// input is normalised to one here. This re-views the same buffer with no copy.
// Only a SharedArrayBuffer, which our inputs never are, falls back to copying.
function arrayBufferBacked(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
  return bytes.buffer instanceof ArrayBuffer
    ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    : new Uint8Array(bytes);
}
