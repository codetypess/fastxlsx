import type { DataValidation, Hyperlink, SetDataValidationOptions } from "./types.js";
import { XlsxError } from "./errors.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "./utils/xml-read.js";
import { decodeXmlText, escapeRegex, escapeXmlText, parseAttributes, serializeAttributes } from "./utils/xml.js";

export const HYPERLINK_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const AUTO_FILTER_FOLLOWING_TAGS = [
  "sortState",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

const DATA_VALIDATIONS_FOLLOWING_TAGS = [
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

export function parseSheetAutoFilter(sheetXml: string): string | null {
  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");
  if (!autoFilterTag) {
    return null;
  }

  const ref = getTagAttr(autoFilterTag, "ref");
  return ref ? normalizeRangeRef(ref) : null;
}

export function upsertAutoFilterInSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeRangeRef(range);
  const autoFilterXml = `<autoFilter ref="${normalizedRange}"/>`;
  const autoFilterTag = findFirstXmlTag(sheetXml, "autoFilter");

  if (autoFilterTag) {
    return replaceXmlTagSource(sheetXml, autoFilterTag, autoFilterXml);
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, AUTO_FILTER_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + autoFilterXml + sheetXml.slice(insertionIndex);
}

export function removeAutoFilterFromSheetXml(sheetXml: string): string {
  let nextSheetXml = sheetXml;

  const autoFilterTag = findFirstXmlTag(nextSheetXml, "autoFilter");
  if (autoFilterTag) {
    nextSheetXml = replaceXmlTagSource(nextSheetXml, autoFilterTag, "");
  }

  const sortStateTag = findFirstXmlTag(nextSheetXml, "sortState");
  if (sortStateTag) {
    nextSheetXml = replaceXmlTagSource(nextSheetXml, sortStateTag, "");
  }

  return nextSheetXml;
}

export function parseSheetDataValidations(sheetXml: string): DataValidation[] {
  const dataValidationsTag = findFirstXmlTag(sheetXml, "dataValidations");
  if (!dataValidationsTag || dataValidationsTag.innerXml === null) {
    return [];
  }

  return parseDataValidationEntries(dataValidationsTag.innerXml)
    .map((validationTag) => {
      const sqref = getTagAttr(validationTag, "sqref");
      if (!sqref) {
        return null;
      }

      const errorTitle = getTagAttr(validationTag, "errorTitle");
      const error = getTagAttr(validationTag, "error");
      const promptTitle = getTagAttr(validationTag, "promptTitle");
      const prompt = getTagAttr(validationTag, "prompt");
      const formula1 = findFirstXmlTag(validationTag.innerXml ?? "", "formula1")?.innerXml;
      const formula2 = findFirstXmlTag(validationTag.innerXml ?? "", "formula2")?.innerXml;

      return {
        range: normalizeSqref(sqref),
        type: getTagAttr(validationTag, "type") ?? null,
        operator: getTagAttr(validationTag, "operator") ?? null,
        allowBlank: parseOptionalXmlBoolean(getTagAttr(validationTag, "allowBlank")),
        showInputMessage: parseOptionalXmlBoolean(getTagAttr(validationTag, "showInputMessage")),
        showErrorMessage: parseOptionalXmlBoolean(getTagAttr(validationTag, "showErrorMessage")),
        showDropDown: parseOptionalXmlBoolean(getTagAttr(validationTag, "showDropDown")),
        errorStyle: getTagAttr(validationTag, "errorStyle") ?? null,
        errorTitle: errorTitle ? decodeXmlText(errorTitle) : null,
        error: error ? decodeXmlText(error) : null,
        promptTitle: promptTitle ? decodeXmlText(promptTitle) : null,
        prompt: prompt ? decodeXmlText(prompt) : null,
        imeMode: getTagAttr(validationTag, "imeMode") ?? null,
        formula1: formula1 ? decodeXmlText(formula1) : null,
        formula2: formula2 ? decodeXmlText(formula2) : null,
      };
    })
    .filter((validation): validation is DataValidation => validation !== null);
}

export function buildDataValidationXml(range: string, options: SetDataValidationOptions): string {
  const attributes: Array<[string, string]> = [["sqref", normalizeSqref(range)]];
  appendOptionalAttribute(attributes, "type", options.type);
  appendOptionalAttribute(attributes, "operator", options.operator);
  appendOptionalBooleanAttribute(attributes, "allowBlank", options.allowBlank);
  appendOptionalBooleanAttribute(attributes, "showInputMessage", options.showInputMessage);
  appendOptionalBooleanAttribute(attributes, "showErrorMessage", options.showErrorMessage);
  appendOptionalBooleanAttribute(attributes, "showDropDown", options.showDropDown);
  appendOptionalAttribute(attributes, "errorStyle", options.errorStyle);
  appendOptionalAttribute(attributes, "errorTitle", options.errorTitle);
  appendOptionalAttribute(attributes, "error", options.error);
  appendOptionalAttribute(attributes, "promptTitle", options.promptTitle);
  appendOptionalAttribute(attributes, "prompt", options.prompt);
  appendOptionalAttribute(attributes, "imeMode", options.imeMode);

  const formulas: string[] = [];
  if (options.formula1 !== undefined) {
    formulas.push(`<formula1>${escapeXmlText(options.formula1)}</formula1>`);
  }
  if (options.formula2 !== undefined) {
    formulas.push(`<formula2>${escapeXmlText(options.formula2)}</formula2>`);
  }

  return formulas.length === 0
    ? `<dataValidation ${serializeAttributes(attributes)}/>`
    : `<dataValidation ${serializeAttributes(attributes)}>${formulas.join("")}</dataValidation>`;
}

export function upsertDataValidationInSheetXml(sheetXml: string, dataValidationXml: string, range: string): string {
  const normalizedRange = normalizeSqref(range);
  const dataValidationsTag = findFirstXmlTag(sheetXml, "dataValidations");
  const dataValidations = ((dataValidationsTag?.innerXml ?? "")
    ? parseDataValidationEntries(dataValidationsTag?.innerXml ?? "").map((validationTag) => ({
        range: normalizeSqref(getTagAttr(validationTag, "sqref") ?? ""),
        xml: validationTag.source,
      }))
    : []
  ).filter((validation) => validation.range !== normalizedRange);

  dataValidations.push({ range: normalizedRange, xml: dataValidationXml });
  const nextDataValidationsXml = buildCountedXmlContainer(
    "dataValidations",
    dataValidationsTag?.attributesSource ?? "",
    "count",
    dataValidations.map((validation) => validation.xml),
  );

  if (dataValidationsTag) {
    return replaceXmlTagSource(sheetXml, dataValidationsTag, nextDataValidationsXml);
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, DATA_VALIDATIONS_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + nextDataValidationsXml + sheetXml.slice(insertionIndex);
}

export function removeDataValidationFromSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeSqref(range);
  const dataValidationsTag = findFirstXmlTag(sheetXml, "dataValidations");
  if (!dataValidationsTag || dataValidationsTag.innerXml === null) {
    return sheetXml;
  }

  const keptDataValidations = parseDataValidationEntries(dataValidationsTag.innerXml).filter(
    (validationTag) => normalizeSqref(getTagAttr(validationTag, "sqref") ?? "") !== normalizedRange,
  );

  const nextDataValidationsXml =
    keptDataValidations.length === 0
      ? ""
      : buildCountedXmlContainer(
          "dataValidations",
          dataValidationsTag.attributesSource,
          "count",
          keptDataValidations.map((validationTag) => validationTag.source),
        );

  return replaceXmlTagSource(sheetXml, dataValidationsTag, nextDataValidationsXml);
}

export function parseSheetHyperlinks(
  sheetXml: string,
  relationshipTargets: Map<string, string>,
): Hyperlink[] {
  return findXmlTags(sheetXml, "hyperlink").map((tag) => {
    const address = getTagAttr(tag, "ref");
    const relationshipId = getTagAttr(tag, "r:id");
    const location = getTagAttr(tag, "location");
    const tooltip = getTagAttr(tag, "tooltip") ?? null;

    if (!address) {
      return null;
    }

    if (relationshipId) {
      const target = relationshipTargets.get(relationshipId);
      if (!target) {
        return null;
      }

      return {
        address: normalizeCellAddress(address),
        target,
        tooltip,
        type: "external" as const,
      };
    }

    if (!location) {
      return null;
    }

    return {
      address: normalizeCellAddress(address),
      target: location,
      tooltip,
      type: "internal" as const,
    };
  }).filter((hyperlink): hyperlink is Hyperlink => hyperlink !== null);
}

export function parseHyperlinkRelationshipTargets(relationshipsXml: string): Map<string, string> {
  const targets = new Map<string, string>();

  for (const relationshipTag of findXmlTags(relationshipsXml, "Relationship")) {
    if (!relationshipTag.selfClosing) {
      continue;
    }

    const relationshipId = getTagAttr(relationshipTag, "Id");
    const type = getTagAttr(relationshipTag, "Type");
    const target = getTagAttr(relationshipTag, "Target");

    if (!relationshipId || !type || !target || type !== HYPERLINK_RELATIONSHIP_TYPE) {
      continue;
    }

    targets.set(relationshipId, decodeXmlText(target));
  }

  return targets;
}

export function getHyperlinkRelationshipId(sheetXml: string, address: string): string | null {
  const normalizedAddress = normalizeCellAddress(address);

  for (const hyperlinkTag of findXmlTags(sheetXml, "hyperlink")) {
    const ref = getTagAttr(hyperlinkTag, "ref");
    if (!ref || normalizeCellAddress(ref) !== normalizedAddress) {
      continue;
    }

    return getTagAttr(hyperlinkTag, "r:id") ?? null;
  }

  return null;
}

export function buildInternalHyperlinkXml(address: string, location: string, tooltip?: string): string {
  const attributes: Array<[string, string]> = [["ref", address], ["location", location]];
  if (tooltip) {
    attributes.push(["tooltip", tooltip]);
  }

  return `<hyperlink ${serializeAttributes(attributes)}/>`;
}

export function buildExternalHyperlinkXml(address: string, relationshipId: string, tooltip?: string): string {
  const attributes: Array<[string, string]> = [["ref", address], ["r:id", relationshipId]];
  if (tooltip) {
    attributes.push(["tooltip", tooltip]);
  }

  return `<hyperlink ${serializeAttributes(attributes)}/>`;
}

export function upsertHyperlinkInSheetXml(sheetXml: string, hyperlinkXml: string, address: string): string {
  const normalizedAddress = normalizeCellAddress(address);
  const hyperlinksTag = findFirstXmlTag(sheetXml, "hyperlinks");
  const hyperlinksInnerXml = hyperlinksTag?.innerXml ?? "";

  const hyperlinks = (hyperlinksInnerXml
    ? findXmlTags(hyperlinksInnerXml, "hyperlink").map((tag) => {
        const ref = getTagAttr(tag, "ref");
        return {
          address: ref ? normalizeCellAddress(ref) : "",
          xml: tag.source,
        };
      })
    : []
  ).filter((hyperlink) => hyperlink.address !== normalizedAddress);
  hyperlinks.push({ address: normalizedAddress, xml: hyperlinkXml });
  hyperlinks.sort((left, right) => compareCellAddresses(left.address, right.address));

  const nextHyperlinksXml = `<hyperlinks>${hyperlinks.map((hyperlink) => hyperlink.xml).join("")}</hyperlinks>`;

  if (hyperlinksTag) {
    return replaceXmlTagSource(sheetXml, hyperlinksTag, nextHyperlinksXml);
  }

  const closingTag = "</worksheet>";
  const insertionIndex = sheetXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return sheetXml.slice(0, insertionIndex) + nextHyperlinksXml + sheetXml.slice(insertionIndex);
}

export function removeHyperlinkFromSheetXml(sheetXml: string, address: string): string {
  const normalizedAddress = normalizeCellAddress(address);
  const hyperlinksTag = findFirstXmlTag(sheetXml, "hyperlinks");
  if (!hyperlinksTag) {
    return sheetXml;
  }

  const keptHyperlinks = findXmlTags(hyperlinksTag.innerXml ?? "", "hyperlink")
    .map((tag) => {
      const ref = getTagAttr(tag, "ref");
      return {
        address: ref ? normalizeCellAddress(ref) : "",
        xml: tag.source,
      };
    })
    .filter((hyperlink) => hyperlink.address !== normalizedAddress);

  const nextHyperlinksXml =
    keptHyperlinks.length === 0
      ? ""
      : `<hyperlinks>${keptHyperlinks.map((hyperlink) => hyperlink.xml).join("")}</hyperlinks>`;

  return replaceXmlTagSource(sheetXml, hyperlinksTag, nextHyperlinksXml);
}

function replaceXmlTagSource(xml: string, tag: XmlTag, nextSource: string): string {
  return xml.slice(0, tag.start) + nextSource + xml.slice(tag.end);
}

function buildCountedXmlContainer(
  tagName: string,
  attributesSource: string,
  countAttributeName: string,
  childXml: string[],
): string {
  const attributes = parseAttributes(attributesSource);
  const nextAttributes = [...attributes];
  const countIndex = nextAttributes.findIndex(([name]) => name === countAttributeName);

  if (countIndex === -1) {
    nextAttributes.push([countAttributeName, String(childXml.length)]);
  } else {
    nextAttributes[countIndex] = [countAttributeName, String(childXml.length)];
  }

  const serializedAttributes = serializeAttributes(nextAttributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${childXml.join("")}</${tagName}>`;
}

function parseDataValidationEntries(innerXml: string): XmlTag[] {
  return findXmlTags(innerXml, "dataValidation");
}

function findWorksheetChildInsertionIndex(sheetXml: string, followingTagNames: string[]): number {
  let insertionIndex = -1;

  for (const tagName of followingTagNames) {
    const match = sheetXml.match(new RegExp(`<${escapeRegex(tagName)}\\b`));
    if (!match || match.index === undefined) {
      continue;
    }

    if (insertionIndex === -1 || match.index < insertionIndex) {
      insertionIndex = match.index;
    }
  }

  if (insertionIndex !== -1) {
    return insertionIndex;
  }

  const closingTag = "</worksheet>";
  const closingTagIndex = sheetXml.indexOf(closingTag);
  if (closingTagIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return closingTagIndex;
}

function appendOptionalAttribute(attributes: Array<[string, string]>, name: string, value: string | undefined): void {
  if (value !== undefined) {
    attributes.push([name, value]);
  }
}

function appendOptionalBooleanAttribute(attributes: Array<[string, string]>, name: string, value: boolean | undefined): void {
  if (value !== undefined) {
    attributes.push([name, value ? "1" : "0"]);
  }
}

function parseOptionalXmlBoolean(value: string | undefined): boolean | null {
  if (value === undefined) {
    return null;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function compareCellAddresses(left: string, right: string): number {
  const leftCell = splitCellAddress(left);
  const rightCell = splitCellAddress(right);
  return leftCell.rowNumber - rightCell.rowNumber || leftCell.columnNumber - rightCell.columnNumber;
}

function normalizeCellAddress(address: string): string {
  assertCellAddress(address);
  return address.toUpperCase();
}

function normalizeRangeRef(range: string): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  return formatRangeRef(startRow, startColumn, endRow, endColumn);
}

function normalizeSqref(rangeList: string): string {
  const ranges = rangeList
    .trim()
    .split(/\s+/)
    .filter((range) => range.length > 0)
    .map((range) => normalizeRangeRef(range));

  if (ranges.length === 0) {
    throw new XlsxError(`Invalid sqref: ${rangeList}`);
  }

  return ranges.join(" ");
}

function parseRangeRef(range: string): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} {
  const normalizedRange = range.toUpperCase();
  const [startAddress, endAddress = normalizedRange] = normalizedRange.split(":");

  if (!startAddress || !endAddress) {
    throw new XlsxError(`Invalid range reference: ${range}`);
  }

  const start = splitCellAddress(startAddress);
  const end = splitCellAddress(endAddress);

  return {
    startRow: Math.min(start.rowNumber, end.rowNumber),
    endRow: Math.max(start.rowNumber, end.rowNumber),
    startColumn: Math.min(start.columnNumber, end.columnNumber),
    endColumn: Math.max(start.columnNumber, end.columnNumber),
  };
}

function splitCellAddress(address: string): { rowNumber: number; columnNumber: number } {
  assertCellAddress(address);
  const match = address.toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return {
    columnNumber: columnLabelToNumber(match[1]),
    rowNumber: Number(match[2]),
  };
}

function assertCellAddress(address: string): void {
  if (!/^[A-Z]+[1-9]\d*$/i.test(address)) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }
}

function makeCellAddress(rowNumber: number, columnNumber: number): string {
  return `${numberToColumnLabel(columnNumber)}${rowNumber}`;
}

function formatRangeRef(
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): string {
  const startAddress = makeCellAddress(startRow, startColumn);
  const endAddress = makeCellAddress(endRow, endColumn);
  return startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`;
}

function columnLabelToNumber(label: string): number {
  let value = 0;

  for (const character of label) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }

  return value;
}

function numberToColumnLabel(columnNumber: number): string {
  if (!Number.isInteger(columnNumber) || columnNumber < 1) {
    throw new XlsxError(`Invalid column number: ${columnNumber}`);
  }

  let remaining = columnNumber;
  let label = "";

  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return label;
}
