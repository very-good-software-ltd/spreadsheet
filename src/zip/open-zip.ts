import type { ByteRange } from "../io/byte-range";
import { NativeZipArchive } from "./native-zip-archive";
import type { ZipArchive } from "./zip-archive";

export function openZip(source: ByteRange): Promise<ZipArchive> {
  return NativeZipArchive.open(source);
}
