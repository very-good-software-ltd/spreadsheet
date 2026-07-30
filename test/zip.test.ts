import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { FflateZipArchive } from "../src/zip/fflate-zip-archive";

describe("FflateZipArchive", () => {
  it("lists entries with their paths and sizes", () => {
    const bytes = zipSync({
      "a.txt": strToU8("hello"),
      "xl/worksheets/sheet1.xml": strToU8("<worksheet/>"),
    });

    const archive = new FflateZipArchive(bytes);

    expect(
      archive
        .entries()
        .map((entry) => entry.path)
        .sort(),
    ).toEqual(["a.txt", "xl/worksheets/sheet1.xml"]);
    expect(archive.has("a.txt")).toBe(true);
    expect(archive.has("missing")).toBe(false);
  });

  it("reads an entry as decompressed bytes", async () => {
    const bytes = zipSync({ "a.txt": strToU8("hello") });

    const data = await new FflateZipArchive(bytes).read("a.txt");

    expect(new TextDecoder().decode(data)).toBe("hello");
  });

  it("opens an entry as a stream", async () => {
    const bytes = zipSync({ "a.txt": strToU8("hello") });

    const stream = new FflateZipArchive(bytes).openStream("a.txt");
    const collected = await new Response(stream).text();

    expect(collected).toBe("hello");
  });

  it("throws for a missing entry", () => {
    const archive = new FflateZipArchive(zipSync({ "a.txt": strToU8("hello") }));

    expect(() => archive.openStream("missing")).toThrow("Zip entry not found: missing");
  });
});
