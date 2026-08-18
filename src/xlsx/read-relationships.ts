import { readPart } from "../read-part";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";

const Element = {
  Relationship: "Relationship",
} as const;

const Attribute = {
  Id: "Id",
  Type: "Type",
  Target: "Target",
} as const;

const WORKBOOK_PART = "xl/workbook.xml";

export interface Relationship {
  readonly id: string;

  /** The full relationship type URI, which says what the target is for. */
  readonly type: string;

  /** The target as a path from the root of the archive. */
  readonly target: string;
}

/**
 * Everything the part at `partPath` points at, with each target resolved from the
 * root of the archive rather than left relative to the part.
 *
 * Empty when the part has no relationships of its own, which is the ordinary case
 * for a worksheet with no tables, drawings or comments.
 */
export async function readRelationships(
  archive: ZipArchive,
  xml: XmlReader,
  partPath: string,
): Promise<readonly Relationship[]> {
  const relationships: Relationship[] = [];
  const path = relationshipsPathFor(partPath);

  if (!archive.has(path)) {
    return relationships;
  }

  for await (const batch of readPart(archive, xml, path)) {
    for (const event of batch) {
      if (event.type !== "open" || event.name !== Element.Relationship) {
        continue;
      }

      const id = event.attributes[Attribute.Id];
      const target = event.attributes[Attribute.Target];
      if (id !== undefined && target !== undefined) {
        relationships.push({
          id,
          type: event.attributes[Attribute.Type] ?? "",
          target: resolveAgainst(partPath, target),
        });
      }
    }
  }

  return relationships;
}

export async function readWorkbookRelationships(archive: ZipArchive, xml: XmlReader): Promise<Map<string, string>> {
  const relationships = await readRelationships(archive, xml, WORKBOOK_PART);

  return new Map(relationships.map((relationship) => [relationship.id, relationship.target]));
}

// A part's relationships live beside it, in a `_rels` folder, under the part's own
// name with `.rels` on the end.
function relationshipsPathFor(partPath: string): string {
  const cut = partPath.lastIndexOf("/");
  const folder = cut < 0 ? "" : partPath.slice(0, cut + 1);
  const file = partPath.slice(cut + 1);

  return `${folder}_rels/${file}.rels`;
}

// A target starting with a slash is already from the root. Anything else is
// relative to the folder the part is in, and a worksheet reaches a table by
// climbing out of the worksheets folder first.
function resolveAgainst(partPath: string, target: string): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }

  const segments = partPath.split("/").slice(0, -1);

  for (const step of target.split("/")) {
    if (step === "..") {
      segments.pop();
    } else if (step !== "." && step !== "") {
      segments.push(step);
    }
  }

  return segments.join("/");
}
