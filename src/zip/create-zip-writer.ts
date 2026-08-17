import { NativeZipWriter } from "./native-zip-writer";
import type { ZipWriter } from "./zip-writer";

export function createZipWriter(): ZipWriter {
  return new NativeZipWriter();
}
