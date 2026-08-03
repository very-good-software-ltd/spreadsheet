import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { BytesByteRange } from "../src/io/byte-range";
import { openZip } from "../src/zip/open-zip";

// fflate writes the archives and our native reader reads them back, so a mature
// implementation produces the input rather than a hand-authored one.

function eocdOffset(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let pos = bytes.length - 22; pos >= 0; pos--) {
    if (view.getUint32(pos, true) === 0x0605_4b50) {
      return pos;
    }
  }
  throw new Error("fixture has no end of central directory");
}

function centralDirectoryOffset(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return view.getUint32(eocdOffset(bytes) + 16, true);
}

async function collectChunks(stream: ReadableStream<Uint8Array>): Promise<Uint8Array[]> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    chunks.push(value);
  }
  return chunks;
}

function concat(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

function openBytes(bytes: Uint8Array) {
  return openZip(new BytesByteRange(bytes));
}

describe("NativeZipArchive", () => {
  it("reads every entry back to the bytes fflate compressed", async () => {
    const files = {
      "a.txt": strToU8("hello"),
      "empty.bin": new Uint8Array(0),
      "nested/deep/path.xml": strToU8("<r/>"),
      "unïcode-nàme.xml": strToU8("<x/>"),
      "xl/worksheets/sheet1.xml": strToU8(`<worksheet>${"data ".repeat(3000)}</worksheet>`),
    };
    const archive = await openBytes(zipSync(files));

    expect(
      archive
        .entries()
        .map((entry) => entry.path)
        .sort(),
    ).toEqual(Object.keys(files).sort());
    for (const [path, expected] of Object.entries(files)) {
      expect(archive.has(path)).toBe(true);
      expect(await archive.read(path)).toEqual(expected);
    }
  });

  it("reads stored (uncompressed) entries", async () => {
    const files = { "a.txt": strToU8("hello"), "b.bin": strToU8("x".repeat(1000)) };
    const archive = await openBytes(zipSync(files, { level: 0 }));

    for (const [path, expected] of Object.entries(files)) {
      expect(await archive.read(path)).toEqual(expected);
    }
  });

  it("streams a large entry in multiple chunks", async () => {
    const original = strToU8("lorem ipsum ".repeat(100_000));
    const archive = await openBytes(zipSync({ "big.xml": original }));

    const chunks = await collectChunks(archive.openStream("big.xml"));

    expect(chunks.length).toBeGreaterThan(1);
    expect(concat(chunks)).toEqual(original);
  });

  it("reads an entry larger than the read chunk size", async () => {
    const original = strToU8("x".repeat(2_500_000));
    const archive = await openBytes(zipSync({ "big.bin": original }, { level: 0 }));

    expect(await archive.read("big.bin")).toEqual(original);
  });

  it("throws for a missing entry", async () => {
    const archive = await openBytes(zipSync({ "a.txt": strToU8("hi") }));

    expect(() => archive.openStream("missing")).toThrow("Zip entry not found: missing");
  });

  it("reads an archive that has a trailing comment", async () => {
    const base = zipSync({ "a.txt": strToU8("hello") });
    const comment = strToU8("a trailing comment");
    const bytes = new Uint8Array(base.length + comment.length);
    bytes.set(base);
    bytes.set(comment, base.length);
    new DataView(bytes.buffer).setUint16(eocdOffset(base) + 20, comment.length, true);

    const archive = await openBytes(bytes);

    expect(await archive.read("a.txt")).toEqual(strToU8("hello"));
  });

  it("throws on an unsupported compression method", async () => {
    const bytes = zipSync({ "a.txt": strToU8("hello ".repeat(100)) });
    new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint16(
      centralDirectoryOffset(bytes) + 10,
      99,
      true,
    );
    const archive = await openBytes(bytes);

    expect(() => archive.openStream("a.txt")).toThrow("Unsupported compression method 99");
  });

  it("rejects zip64 archives", async () => {
    const bytes = zipSync({ "a.txt": strToU8("hello") });
    new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(
      centralDirectoryOffset(bytes) + 20,
      0xffff_ffff,
      true,
    );

    await expect(openBytes(bytes)).rejects.toThrow("Zip64 archives are not supported");
  });

  it("rejects a buffer too short to be a zip", async () => {
    await expect(openBytes(new Uint8Array([1, 2, 3]))).rejects.toThrow("Malformed zip");
  });
});
