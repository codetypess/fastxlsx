import { XlsxError } from "../errors.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "../utils/xml-read.js";
import { escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";
import { replaceXmlTagSource } from "./workbook-xml.js";

export function getRequiredXmlContainerTag(
  xml: string,
  tagName: string,
  fileLabel: string,
): XmlTag & { innerXml: string } {
  const tag = findFirstXmlTag(xml, tagName);
  if (!tag || tag.innerXml === null) {
    throw new XlsxError(`${fileLabel} is missing <${tagName}>`);
  }

  return tag as XmlTag & { innerXml: string };
}

export function appendFontToStylesXml(stylesXml: string, fontXml: string): string {
  return appendXmlToStylesContainer(stylesXml, "fonts", fontXml);
}

export function appendFillToStylesXml(stylesXml: string, fillXml: string): string {
  return appendXmlToStylesContainer(stylesXml, "fills", fillXml);
}

export function appendBorderToStylesXml(stylesXml: string, borderXml: string): string {
  return appendXmlToStylesContainer(stylesXml, "borders", borderXml);
}

export function replaceFontInStylesXml(stylesXml: string, fontId: number, fontXml: string): string {
  return replaceIndexedXmlInStylesContainer(stylesXml, "fonts", fontId, fontXml, `Font not found: ${fontId}`);
}

export function replaceFillInStylesXml(stylesXml: string, fillId: number, fillXml: string): string {
  return replaceIndexedXmlInStylesContainer(stylesXml, "fills", fillId, fillXml, `Fill not found: ${fillId}`);
}

export function replaceBorderInStylesXml(stylesXml: string, borderId: number, borderXml: string): string {
  return replaceIndexedXmlInStylesContainer(
    stylesXml,
    "borders",
    borderId,
    borderXml,
    `Border not found: ${borderId}`,
  );
}

export function upsertNumberFormatInStylesXml(stylesXml: string, numFmtId: number, formatCode: string): string {
  const numFmtXml = `<numFmt numFmtId="${numFmtId}" formatCode="${escapeXmlText(formatCode)}"/>`;
  const numberFormatsTag = findFirstXmlTag(stylesXml, "numFmts");

  if (!numberFormatsTag || numberFormatsTag.innerXml === null) {
    const fontsTag = findFirstXmlTag(stylesXml, "fonts");
    if (!fontsTag) {
      throw new XlsxError("styles.xml is missing <fonts>");
    }

    return (
      stylesXml.slice(0, fontsTag.start) +
      `<numFmts count="1">${numFmtXml}</numFmts>` +
      stylesXml.slice(fontsTag.start)
    );
  }

  const numFmtTags = findXmlTags(numberFormatsTag.innerXml, "numFmt");
  const matchingTag = numFmtTags.find((tag) => Number(getTagAttr(tag, "numFmtId")) === numFmtId);
  if (matchingTag) {
    const nextInnerXml =
      numberFormatsTag.innerXml.slice(0, matchingTag.start) +
      numFmtXml +
      numberFormatsTag.innerXml.slice(matchingTag.end);
    return replaceXmlTagSource(
      stylesXml,
      numberFormatsTag,
      buildStylesContainerXml("numFmts", numberFormatsTag.attributesSource, nextInnerXml),
    );
  }

  return appendXmlToStylesContainer(stylesXml, "numFmts", numFmtXml);
}

export function appendCellXfToStylesXml(stylesXml: string, xfXml: string): string {
  return appendXmlToStylesContainer(stylesXml, "cellXfs", xfXml);
}

export function replaceCellXfInStylesXml(stylesXml: string, styleId: number, xfXml: string): string {
  return replaceIndexedXmlInStylesContainer(stylesXml, "cellXfs", styleId, xfXml, `Style not found: ${styleId}`);
}

function buildStylesContainerXml(tagName: string, attributesSource: string, innerXml: string): string {
  const attributes = parseAttributes(attributesSource);
  const nextCount = findXmlTags(innerXml, getStylesContainerChildTagName(tagName)).length;
  const nextAttributes = upsertAttribute(attributes, "count", String(nextCount));
  const serializedAttributes = serializeAttributes(nextAttributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</${tagName}>`;
}

function getStylesContainerChildTagName(tagName: string): string {
  if (tagName === "borders") {
    return "border";
  }
  if (tagName === "fonts") {
    return "font";
  }
  if (tagName === "fills") {
    return "fill";
  }
  if (tagName === "cellXfs") {
    return "xf";
  }
  if (tagName === "numFmts") {
    return "numFmt";
  }

  throw new XlsxError(`Unsupported styles container: <${tagName}>`);
}

function appendXmlToStylesContainer(stylesXml: string, containerTagName: string, childXml: string): string {
  const containerTag = getRequiredXmlContainerTag(stylesXml, containerTagName, "styles.xml");
  const innerXml = containerTag.innerXml;
  const trailingWhitespace = innerXml.match(/\s*$/)?.[0] ?? "";
  const innerXmlWithoutTrailing = innerXml.slice(0, innerXml.length - trailingWhitespace.length);
  const closingIndentMatch = trailingWhitespace.match(/\n([ \t]*)$/);
  const entryPrefix = closingIndentMatch ? `\n${closingIndentMatch[1]}  ` : "";
  const nextInnerXml = `${innerXmlWithoutTrailing}${entryPrefix}${childXml}${trailingWhitespace}`;
  const nextContainerXml = buildStylesContainerXml(containerTagName, containerTag.attributesSource, nextInnerXml);

  return replaceXmlTagSource(stylesXml, containerTag, nextContainerXml);
}

function replaceIndexedXmlInStylesContainer(
  stylesXml: string,
  containerTagName: string,
  targetIndex: number,
  childXml: string,
  missingMessage: string,
): string {
  const containerTag = getRequiredXmlContainerTag(stylesXml, containerTagName, "styles.xml");
  const childTags = findXmlTags(containerTag.innerXml, getStylesContainerChildTagName(containerTagName));
  const childTag = childTags[targetIndex];
  if (!childTag) {
    throw new XlsxError(missingMessage);
  }

  const nextInnerXml =
    containerTag.innerXml.slice(0, childTag.start) + childXml + containerTag.innerXml.slice(childTag.end);
  const nextContainerXml = buildStylesContainerXml(containerTagName, containerTag.attributesSource, nextInnerXml);

  return replaceXmlTagSource(stylesXml, containerTag, nextContainerXml);
}

function upsertAttribute(
  attributes: Array<[string, string]>,
  name: string,
  value: string | null,
): Array<[string, string]> {
  const nextAttributes: Array<[string, string]> = [];
  let found = false;

  for (const [attributeName, attributeValue] of attributes) {
    if (attributeName !== name) {
      nextAttributes.push([attributeName, attributeValue]);
      continue;
    }

    found = true;
    if (value !== null) {
      nextAttributes.push([attributeName, value]);
    }
  }

  if (!found && value !== null) {
    nextAttributes.push([name, value]);
  }

  return nextAttributes;
}
