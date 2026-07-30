import { FflateZipArchive } from "./fflate-zip-archive";
import type { ZipArchive } from "./zip-archive";

export function openZip(bytes: Uint8Array): ZipArchive {
  return new FflateZipArchive(bytes);
}
