import { type Unzipped, unzipSync } from "fflate";
import type { ZipArchive, ZipEntry } from "./zip-archive";

export class FflateZipArchive implements ZipArchive {
  private readonly files: Unzipped;
  private readonly entryList: readonly ZipEntry[];

  constructor(bytes: Uint8Array) {
    // unzipSync decompresses every entry up front, so the archive is eager
    // despite the streaming interface.
    this.files = unzipSync(bytes);
    this.entryList = Object.entries(this.files).map(([path, data]) => ({
      path,
      size: data.length,
    }));
  }

  entries(): readonly ZipEntry[] {
    return this.entryList;
  }

  has(path: string): boolean {
    return Object.hasOwn(this.files, path);
  }

  read(path: string): Promise<Uint8Array> {
    return Promise.resolve(this.bytesFor(path));
  }

  openStream(path: string): ReadableStream<Uint8Array> {
    const data = this.bytesFor(path);
    return new ReadableStream<Uint8Array>({
      start(controller): void {
        controller.enqueue(data);
        controller.close();
      },
    });
  }

  private bytesFor(path: string): Uint8Array {
    const data = this.files[path];
    if (data === undefined) {
      throw new Error(`Zip entry not found: ${path}`);
    }
    return data;
  }
}
