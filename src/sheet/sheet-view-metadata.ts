import type { FreezePane, SheetSelection } from "../types.js";
import { XlsxError } from "../errors.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "../utils/xml-read.js";
import { escapeRegex, parseAttributes, serializeAttributes } from "../utils/xml.js";

const SHEET_VIEWS_FOLLOWING_TAGS = [
  "sheetFormatPr",
  "cols",
  "sheetData",
  "autoFilter",
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

export function parseSheetFreezePane(sheetXml: string): FreezePane | null {
  const { sheetViewTag } = getSheetViewTags(sheetXml);
  const paneTag =
    sheetViewTag && sheetViewTag.innerXml !== null ? findFirstXmlTag(sheetViewTag.innerXml, "pane") : null;

  if (!paneTag) {
    return null;
  }

  const state = getTagAttr(paneTag, "state");
  if (state && state !== "frozen" && state !== "frozenSplit") {
    return null;
  }

  const columnCount = Number(getTagAttr(paneTag, "xSplit") ?? "0");
  const rowCount = Number(getTagAttr(paneTag, "ySplit") ?? "0");
  if ((!Number.isFinite(columnCount) || columnCount < 0) && (!Number.isFinite(rowCount) || rowCount < 0)) {
    return null;
  }

  if (columnCount === 0 && rowCount === 0) {
    return null;
  }

  return {
    columnCount: Number.isFinite(columnCount) ? columnCount : 0,
    rowCount: Number.isFinite(rowCount) ? rowCount : 0,
    topLeftCell: getTagAttr(paneTag, "topLeftCell") ?? makeCellAddress(rowCount + 1, columnCount + 1),
    activePane: normalizePaneName(getTagAttr(paneTag, "activePane")),
  };
}

export function parseSheetSelection(sheetXml: string): SheetSelection | null {
  const freezePane = parseSheetFreezePane(sheetXml);
  const selections = parseSheetSelectionEntries(sheetXml);
  if (selections.length === 0) {
    return null;
  }

  const targetPane = freezePane?.activePane ?? null;
  const selection =
    selections.find((candidate) => candidate.pane === targetPane) ??
    selections.find((candidate) => candidate.activeCell !== null || candidate.range !== null) ??
    selections[0];

  return selection ?? null;
}

export function upsertFreezePaneInSheetXml(sheetXml: string, columnCount: number, rowCount: number): string {
  const paneXml = buildFreezePaneXml(columnCount, rowCount);
  const selectionsXml = buildFreezePaneSelectionsXml(columnCount, rowCount);
  const { sheetViewsTag, sheetViewTag } = getSheetViewTags(sheetXml);

  if (!sheetViewsTag) {
    const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, SHEET_VIEWS_FOLLOWING_TAGS);
    return (
      sheetXml.slice(0, insertionIndex) +
      `<sheetViews><sheetView workbookViewId="0">${paneXml}${selectionsXml}</sheetView></sheetViews>` +
      sheetXml.slice(insertionIndex)
    );
  }

  if (!sheetViewTag) {
    return replaceXmlTagSource(
      sheetXml,
      sheetViewsTag,
      `<sheetViews><sheetView workbookViewId="0">${paneXml}${selectionsXml}</sheetView></sheetViews>`,
    );
  }

  const attributes = parseAttributes(sheetViewTag.attributesSource);
  ensureXmlAttribute(attributes, "workbookViewId", "0");

  const innerXml = sheetViewTag.innerXml ?? "";
  const cleanedInnerXml = removeXmlTagsFromInnerXml(innerXml, [
    ...findXmlTags(innerXml, "pane"),
    ...findXmlTags(innerXml, "selection"),
  ]);
  const serializedAttributes = serializeAttributes(attributes);
  const nextSheetViewXml = `<sheetView${serializedAttributes ? ` ${serializedAttributes}` : ""}>${paneXml}${selectionsXml}${cleanedInnerXml}</sheetView>`;

  return replaceNestedXmlTagSource(sheetXml, sheetViewsTag, sheetViewTag, nextSheetViewXml);
}

export function removeFreezePaneFromSheetXml(sheetXml: string): string {
  const { sheetViewsTag, sheetViewTag } = getSheetViewTags(sheetXml);
  if (!sheetViewsTag || !sheetViewTag) {
    return sheetXml;
  }

  const attributes = parseAttributes(sheetViewTag.attributesSource);
  ensureXmlAttribute(attributes, "workbookViewId", "0");

  const innerXml = sheetViewTag.innerXml ?? "";
  const paneTag = findFirstXmlTag(innerXml, "pane");
  if (!paneTag) {
    return sheetXml;
  }

  const activePane = normalizePaneName(getTagAttr(paneTag, "activePane"));
  const selectionTags = findXmlTags(innerXml, "selection");
  const selections = selectionTags.map((tag) => ({
    attributes: parseAttributes(tag.attributesSource),
    xml: tag.source,
  }));
  const preferredSelection =
    selections.find((selection) => selection.attributes.find(([name]) => name === "pane")?.[1] === activePane) ??
    selections.find((selection) => selection.attributes.some(([name]) => name === "activeCell" || name === "sqref")) ??
    selections[0];
  const nextSelectionXml = preferredSelection
    ? buildSelectionXml(preferredSelection.attributes.filter(([name]) => name !== "pane"))
    : "";
  const cleanedInnerXml = removeXmlTagsFromInnerXml(innerXml, [paneTag, ...selectionTags]);
  const serializedAttributes = serializeAttributes(attributes);
  const nextSheetViewXml = `<sheetView${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nextSelectionXml}${cleanedInnerXml}</sheetView>`;

  return replaceNestedXmlTagSource(sheetXml, sheetViewsTag, sheetViewTag, nextSheetViewXml);
}

export function upsertSheetSelectionInSheetXml(
  sheetXml: string,
  activeCell: string,
  range: string,
): string {
  const freezePane = parseSheetFreezePane(sheetXml);
  const targetPane = freezePane?.activePane ?? null;
  const nextSelectionXml = buildSelectionXml([
    ...(targetPane ? [["pane", targetPane] as [string, string]] : []),
    ["activeCell", activeCell],
    ["sqref", range],
  ]);
  const { sheetViewsTag, sheetViewTag } = getSheetViewTags(sheetXml);

  if (!sheetViewsTag) {
    const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, SHEET_VIEWS_FOLLOWING_TAGS);
    return (
      sheetXml.slice(0, insertionIndex) +
      `<sheetViews><sheetView workbookViewId="0">${nextSelectionXml}</sheetView></sheetViews>` +
      sheetXml.slice(insertionIndex)
    );
  }

  if (!sheetViewTag) {
    return replaceXmlTagSource(
      sheetXml,
      sheetViewsTag,
      `<sheetViews><sheetView workbookViewId="0">${nextSelectionXml}</sheetView></sheetViews>`,
    );
  }

  const attributes = parseAttributes(sheetViewTag.attributesSource);
  ensureXmlAttribute(attributes, "workbookViewId", "0");

  const innerXml = sheetViewTag.innerXml ?? "";
  const selectionTags = findXmlTags(innerXml, "selection");
  let replaced = false;
  let cursor = 0;
  let nextInnerXml = "";

  for (const selectionTag of selectionTags) {
    nextInnerXml += innerXml.slice(cursor, selectionTag.start);
    const selectionPane = normalizePaneName(getTagAttr(selectionTag, "pane"));
    const matchesTargetPane = selectionPane === targetPane;

    if (matchesTargetPane || (!replaced && targetPane === null && selectionPane === null)) {
      replaced = true;
      nextInnerXml += nextSelectionXml;
    } else {
      nextInnerXml += selectionTag.source;
    }

    cursor = selectionTag.end;
  }

  nextInnerXml += innerXml.slice(cursor);
  if (!replaced) {
    nextInnerXml += nextSelectionXml;
  }

  const serializedAttributes = serializeAttributes(attributes);
  const nextSheetViewXml = `<sheetView${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nextInnerXml}</sheetView>`;

  return replaceNestedXmlTagSource(sheetXml, sheetViewsTag, sheetViewTag, nextSheetViewXml);
}

function getXmlTagInnerStart(tag: XmlTag): number {
  if (tag.innerXml === null) {
    return tag.end;
  }

  return tag.end - tag.innerXml.length - `</${tag.tagName}>`.length;
}

function replaceXmlTagSource(xml: string, tag: XmlTag, nextSource: string): string {
  return xml.slice(0, tag.start) + nextSource + xml.slice(tag.end);
}

function replaceNestedXmlTagSource(xml: string, parentTag: XmlTag, childTag: XmlTag, nextSource: string): string {
  const parentInnerStart = getXmlTagInnerStart(parentTag);
  return (
    xml.slice(0, parentInnerStart + childTag.start) +
    nextSource +
    xml.slice(parentInnerStart + childTag.end)
  );
}

function removeXmlTagsFromInnerXml(innerXml: string, tags: XmlTag[]): string {
  return [...tags]
    .sort((left, right) => right.start - left.start)
    .reduce(
      (currentXml, tag) => currentXml.slice(0, tag.start) + currentXml.slice(tag.end),
      innerXml,
    );
}

function getSheetViewTags(sheetXml: string): {
  sheetViewTag: XmlTag | null;
  sheetViewsTag: XmlTag | null;
} {
  const sheetViewsTag = findFirstXmlTag(sheetXml, "sheetViews");
  const sheetViewTag =
    sheetViewsTag && sheetViewsTag.innerXml !== null ? findFirstXmlTag(sheetViewsTag.innerXml, "sheetView") : null;

  return { sheetViewTag, sheetViewsTag };
}

function ensureXmlAttribute(attributes: Array<[string, string]>, name: string, value: string): void {
  if (!attributes.some(([candidateName]) => candidateName === name)) {
    attributes.push([name, value]);
  }
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

function buildFreezePaneXml(columnCount: number, rowCount: number): string {
  const attributes: Array<[string, string]> = [["state", "frozen"]];
  if (columnCount > 0) {
    attributes.push(["xSplit", String(columnCount)]);
  }
  if (rowCount > 0) {
    attributes.push(["ySplit", String(rowCount)]);
  }
  attributes.push(["topLeftCell", makeCellAddress(rowCount + 1, columnCount + 1)]);
  const activePane = getFreezePaneActivePane(columnCount, rowCount);
  if (activePane) {
    attributes.push(["activePane", activePane]);
  }

  return `<pane ${serializeAttributes(attributes)}/>`;
}

function buildFreezePaneSelectionsXml(columnCount: number, rowCount: number): string {
  const topLeftCell = makeCellAddress(rowCount + 1, columnCount + 1);

  if (columnCount > 0 && rowCount > 0) {
    return [
      buildSelectionXml([["pane", "topRight"]]),
      buildSelectionXml([["pane", "bottomLeft"]]),
      buildSelectionXml([["pane", "bottomRight"], ["activeCell", topLeftCell], ["sqref", topLeftCell]]),
    ].join("");
  }

  if (columnCount > 0) {
    return buildSelectionXml([["pane", "topRight"], ["activeCell", topLeftCell], ["sqref", topLeftCell]]);
  }

  return buildSelectionXml([["pane", "bottomLeft"], ["activeCell", topLeftCell], ["sqref", topLeftCell]]);
}

function parseSheetSelectionEntries(sheetXml: string): SheetSelection[] {
  const { sheetViewTag } = getSheetViewTags(sheetXml);
  return findXmlTags(sheetViewTag?.innerXml ?? sheetXml, "selection").map((tag) => {
    const activeCell = getTagAttr(tag, "activeCell");
    const sqref = getTagAttr(tag, "sqref");

    return {
      activeCell: activeCell ? normalizeCellAddress(activeCell) : null,
      range: sqref ? normalizeSqref(sqref) : null,
      pane: normalizePaneName(getTagAttr(tag, "pane")),
    };
  });
}

function buildSelectionXml(attributes: Array<[string, string]>): string {
  return attributes.length === 0 ? "<selection/>" : `<selection ${serializeAttributes(attributes)}/>`;
}

function getFreezePaneActivePane(
  columnCount: number,
  rowCount: number,
): "bottomLeft" | "topRight" | "bottomRight" | null {
  if (columnCount > 0 && rowCount > 0) {
    return "bottomRight";
  }

  if (columnCount > 0) {
    return "topRight";
  }

  if (rowCount > 0) {
    return "bottomLeft";
  }

  return null;
}

function normalizePaneName(
  value: string | undefined,
): "bottomLeft" | "topRight" | "bottomRight" | null {
  if (value === "bottomLeft" || value === "topRight" || value === "bottomRight") {
    return value;
  }

  return null;
}

function normalizeCellAddress(address: string): string {
  assertCellAddress(address);
  return address.toUpperCase();
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

function normalizeRangeRef(range: string): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  return formatRangeRef(startRow, startColumn, endRow, endColumn);
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
