import { writeXmlEvent, XML_DECLARATION } from "../xml/write-xml";
import type { XmlEvent, XmlOpen } from "../xml/xml-reader";

const Element = {
  Workbook: "workbook",
  CalculationProperties: "calcPr",
  Override: "Override",
  Relationship: "Relationship",
} as const;

const Attribute = {
  FullCalculationOnLoad: "fullCalcOnLoad",
  PartName: "PartName",
  Target: "Target",
} as const;

// Elements the schema places after calcPr inside the workbook part. Inserting a
// calcPr before the first of these keeps the sequence valid, and a workbook with
// none of them takes it just before the closing tag.
const AFTER_CALCULATION_PROPERTIES: ReadonlySet<string> = new Set([
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst",
]);

/**
 * Rewrites the workbook part so a spreadsheet application recalculates every
 * formula when it opens the file.
 *
 * Changing a value leaves the cached result of every formula that depends on it
 * stale, and those cached results are what gets displayed until something
 * recalculates. This is what makes that happen.
 */
export async function* withRecalculationOnLoad(events: AsyncIterable<readonly XmlEvent[]>): AsyncIterable<string> {
  yield XML_DECLARATION;

  let written = false;
  let depth = 0;

  for await (const batch of events) {
    for (const event of batch) {
      if (event.type === "open") {
        depth += 1;

        if (event.name === Element.CalculationProperties) {
          written = true;
          yield writeXmlEvent(withFullCalculation(event));
          continue;
        }

        if (!written && depth === 2 && AFTER_CALCULATION_PROPERTIES.has(event.name)) {
          written = true;
          yield calculationProperties();
        }

        yield writeXmlEvent(event);
        continue;
      }

      if (event.type === "close") {
        if (!written && event.name === Element.Workbook) {
          written = true;
          yield calculationProperties();
        }
        depth -= 1;
      }

      yield writeXmlEvent(event);
    }
  }
}

function withFullCalculation(event: XmlOpen): XmlEvent {
  return {
    ...event,
    attributes: { ...event.attributes, [Attribute.FullCalculationOnLoad]: "1" },
  };
}

function calculationProperties(): string {
  const attributes = { [Attribute.FullCalculationOnLoad]: "1" };

  return (
    writeXmlEvent({ type: "open", name: Element.CalculationProperties, attributes }) +
    writeXmlEvent({ type: "close", name: Element.CalculationProperties })
  );
}

/** Rewrites the content types part without the override for `partName`. */
export function withoutContentTypeOverride(
  events: AsyncIterable<readonly XmlEvent[]>,
  partName: string,
): AsyncIterable<string> {
  return withoutElements(
    events,
    (event) => event.name === Element.Override && event.attributes[Attribute.PartName] === partName,
  );
}

/** Rewrites a relationships part without the relationship pointing at `target`. */
export function withoutRelationshipTo(
  events: AsyncIterable<readonly XmlEvent[]>,
  target: string,
): AsyncIterable<string> {
  return withoutElements(
    events,
    (event) => event.name === Element.Relationship && event.attributes[Attribute.Target] === target,
  );
}

async function* withoutElements(
  events: AsyncIterable<readonly XmlEvent[]>,
  matches: (event: XmlOpen) => boolean,
): AsyncIterable<string> {
  yield XML_DECLARATION;

  let skipping = 0;

  for await (const batch of events) {
    for (const event of batch) {
      if (skipping > 0) {
        if (event.type === "open") {
          skipping += 1;
        } else if (event.type === "close") {
          skipping -= 1;
        }
        continue;
      }

      if (event.type === "open" && matches(event)) {
        skipping = 1;
        continue;
      }

      yield writeXmlEvent(event);
    }
  }
}
