import { BytesByteRange } from "../io/byte-range";
import { NativeZipArchive } from "./native-zip-archive";
import type { ZipArchive } from "./zip-archive";

export function openZip(bytes: Uint8Array): Promise<ZipArchive> {
  return NativeZipArchive.open(new BytesByteRange(bytes));
}
