import { XlsxError } from "../errors.js";
import { compareRangeRefs, normalizeRangeRef } from "./sheet-address.js";
import { buildCountedXmlContainer, replaceXmlTagSource } from "./sheet-xml.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "../utils/xml-read.js";

export function parseMergedRanges(sheetXml: string): string[] {
  const mergeCellsTag = findFirstXmlTag(sheetXml, "mergeCells");
  if (!mergeCellsTag?.innerXml) {
    return [];
  }

  return findXmlTags(mergeCellsTag.innerXml, "mergeCell")
    .filter((tag) => tag.selfClosing)
    .map((tag) => getTagAttr(tag, "ref"))
    .filter((ref): ref is string => ref !== undefined)
    .map((ref) => normalizeRangeRef(ref));
}

export function updateMergedRanges(sheetXml: string, ranges: string[]): string {
  const normalizedRanges = [...new Set(ranges.map(normalizeRangeRef))].sort(compareRangeRefs);
  const mergeCellsTag = findFirstXmlTag(sheetXml, "mergeCells");

  if (normalizedRanges.length === 0) {
    if (!mergeCellsTag) {
      return sheetXml;
    }

    return replaceXmlTagSource(sheetXml, mergeCellsTag, "");
  }

  const mergeCellsXml = buildCountedXmlContainer(
    "mergeCells",
    mergeCellsTag?.attributesSource ?? "",
    "count",
    normalizedRanges.map((range) => `<mergeCell ref="${range}"/>`),
  );

  if (mergeCellsTag) {
    return replaceXmlTagSource(sheetXml, mergeCellsTag, mergeCellsXml);
  }

  const sheetDataCloseTag = "</sheetData>";
  const insertionIndex = sheetXml.indexOf(sheetDataCloseTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </sheetData>");
  }

  const anchorIndex = insertionIndex + sheetDataCloseTag.length;
  return sheetXml.slice(0, anchorIndex) + mergeCellsXml + sheetXml.slice(anchorIndex);
}
