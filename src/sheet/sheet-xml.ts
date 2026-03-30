import { XlsxError } from "../errors.js";
import { findXmlTags, type XmlTag } from "../utils/xml-read.js";
import { escapeRegex, parseAttributes, serializeAttributes } from "../utils/xml.js";

export function getXmlTagInnerStart(tag: XmlTag): number {
  if (tag.innerXml === null) {
    return tag.end;
  }

  return tag.end - tag.innerXml.length - `</${tag.tagName}>`.length;
}

export function replaceXmlTagSource(
  xml: string,
  tagOrSource: XmlTag | string,
  nextSource: string,
): string {
  if (typeof tagOrSource === "string") {
    const index = xml.indexOf(tagOrSource);
    if (index === -1) {
      return xml;
    }

    return xml.slice(0, index) + nextSource + xml.slice(index + tagOrSource.length);
  }

  return xml.slice(0, tagOrSource.start) + nextSource + xml.slice(tagOrSource.end);
}

export function replaceNestedXmlTagSource(
  xml: string,
  parentTag: XmlTag,
  childTag: XmlTag,
  nextSource: string,
): string {
  const parentInnerStart = getXmlTagInnerStart(parentTag);
  return (
    xml.slice(0, parentInnerStart + childTag.start) +
    nextSource +
    xml.slice(parentInnerStart + childTag.end)
  );
}

export function removeXmlTagsFromInnerXml(innerXml: string, tags: XmlTag[]): string {
  return [...tags]
    .sort((left, right) => right.start - left.start)
    .reduce((currentXml, tag) => currentXml.slice(0, tag.start) + currentXml.slice(tag.end), innerXml);
}

export function rewriteXmlTagsByName(
  xml: string,
  tagName: string,
  rewriteTag: (tag: XmlTag) => string,
): string {
  const tags = findXmlTags(xml, tagName);
  if (tags.length === 0) {
    return xml;
  }

  let nextXml = "";
  let cursor = 0;

  for (const tag of tags) {
    nextXml += xml.slice(cursor, tag.start);
    nextXml += rewriteTag(tag);
    cursor = tag.end;
  }

  nextXml += xml.slice(cursor);
  return nextXml;
}

export function buildCountedXmlContainer(
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

export function buildXmlContainer(tagName: string, attributesSource: string, innerXml: string): string {
  const serializedAttributes = serializeAttributes(parseAttributes(attributesSource));
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</${tagName}>`;
}

export function buildXmlElement(tagName: string, attributes: Array<[string, string]>, innerXml: string): string {
  const serializedAttributes = serializeAttributes(attributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</${tagName}>`;
}

export function buildSelfClosingXmlElement(tagName: string, attributes: Array<[string, string]>): string {
  const serializedAttributes = serializeAttributes(attributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
}

export function findWorksheetChildInsertionIndex(sheetXml: string, followingTagNames: string[]): number {
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
