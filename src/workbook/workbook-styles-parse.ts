import type {
  CellBorderColor,
  CellBorderDefinition,
  CellBorderSideDefinition,
  CellFillColor,
  CellFillDefinition,
  CellFontColor,
  CellFontDefinition,
  CellStyleAlignment,
  CellStyleDefinition,
} from "../types.js";
import { decodeXmlText, getXmlAttr, parseAttributes } from "../utils/xml.js";
import { findFirstXmlTag, findXmlTags, getTagAttr, type XmlTag } from "../utils/xml-read.js";
import { getRequiredXmlContainerTag } from "./workbook-styles-container.js";

export interface ParsedCellStyle {
  alignmentAttributes: Array<[string, string]> | null;
  attributes: Array<[string, string]>;
  definition: CellStyleDefinition;
  extraChildrenXml: string;
}

export interface ParsedFont {
  definition: CellFontDefinition;
  extraChildrenXml: string;
}

export interface ParsedFill {
  definition: CellFillDefinition;
  extraChildrenXml: string;
}

export interface ParsedBorder {
  definition: CellBorderDefinition;
  extraChildrenXml: string;
}

export interface StylesCache {
  borders: ParsedBorder[];
  cellXfs: ParsedCellStyle[];
  fills: ParsedFill[];
  fonts: ParsedFont[];
  numberFormats: Map<number, string>;
  path: string;
  xml: string;
}

export function parseStylesXml(path: string, xml: string): StylesCache {
  const bordersTag = getRequiredXmlContainerTag(xml, "borders", "styles.xml");
  const fontsTag = getRequiredXmlContainerTag(xml, "fonts", "styles.xml");
  const fillsTag = getRequiredXmlContainerTag(xml, "fills", "styles.xml");
  const cellXfsTag = getRequiredXmlContainerTag(xml, "cellXfs", "styles.xml");

  return {
    path,
    xml,
    borders: findXmlTags(bordersTag.innerXml, "border").map((tag) => parseBorder(tag.source)),
    fills: findXmlTags(fillsTag.innerXml, "fill").map((tag) => parseFill(tag.source)),
    fonts: findXmlTags(fontsTag.innerXml, "font").map((tag) => parseFont(tag.source)),
    numberFormats: parseNumberFormats(xml),
    cellXfs: findXmlTags(cellXfsTag.innerXml, "xf").map((tag) => parseCellStyle(tag.attributesSource, tag.innerXml ?? "")),
  };
}

function parseNumberFormats(stylesXml: string): Map<number, string> {
  const numberFormats = new Map<number, string>();
  const numberFormatsTag = findFirstXmlTag(stylesXml, "numFmts");
  if (!numberFormatsTag || numberFormatsTag.innerXml === null) {
    return numberFormats;
  }

  for (const numFmtTag of findXmlTags(numberFormatsTag.innerXml, "numFmt")) {
    const numFmtIdText = getTagAttr(numFmtTag, "numFmtId");
    const formatCode = getTagAttr(numFmtTag, "formatCode");
    if (numFmtIdText === undefined || formatCode === undefined) {
      continue;
    }

    numberFormats.set(Number(numFmtIdText), decodeXmlText(formatCode));
  }

  return numberFormats;
}

function parseBorder(borderXml: string): ParsedBorder {
  const borderTag = findFirstXmlTag(borderXml, "border");
  const borderAttributes = parseAttributes(borderTag?.attributesSource ?? "");
  let remainingXml = borderTag?.innerXml ?? "";

  const [leftTag, remainingAfterLeft] = takeFirstXmlTagByName(remainingXml, "left");
  remainingXml = remainingAfterLeft;
  const [rightTag, remainingAfterRight] = takeFirstXmlTagByName(remainingXml, "right");
  remainingXml = remainingAfterRight;
  const [topTag, remainingAfterTop] = takeFirstXmlTagByName(remainingXml, "top");
  remainingXml = remainingAfterTop;
  const [bottomTag, remainingAfterBottom] = takeFirstXmlTagByName(remainingXml, "bottom");
  remainingXml = remainingAfterBottom;
  const [diagonalTag, remainingAfterDiagonal] = takeFirstXmlTagByName(remainingXml, "diagonal");
  remainingXml = remainingAfterDiagonal;
  const [verticalTag, remainingAfterVertical] = takeFirstXmlTagByName(remainingXml, "vertical");
  remainingXml = remainingAfterVertical;
  const [horizontalTag, remainingAfterHorizontal] = takeFirstXmlTagByName(remainingXml, "horizontal");
  remainingXml = remainingAfterHorizontal;

  return {
    definition: {
      left: parseBorderSideDefinition(leftTag?.source ?? null),
      right: parseBorderSideDefinition(rightTag?.source ?? null),
      top: parseBorderSideDefinition(topTag?.source ?? null),
      bottom: parseBorderSideDefinition(bottomTag?.source ?? null),
      diagonal: parseBorderSideDefinition(diagonalTag?.source ?? null),
      vertical: parseBorderSideDefinition(verticalTag?.source ?? null),
      horizontal: parseBorderSideDefinition(horizontalTag?.source ?? null),
      diagonalUp: parseOptionalBooleanAttribute(borderAttributes, "diagonalUp"),
      diagonalDown: parseOptionalBooleanAttribute(borderAttributes, "diagonalDown"),
      outline: parseOptionalBooleanAttribute(borderAttributes, "outline"),
    },
    extraChildrenXml: /\S/.test(remainingXml) ? remainingXml : "",
  };
}

function parseFill(fillXml: string): ParsedFill {
  const fillTag = findFirstXmlTag(fillXml, "fill");
  const innerXml = fillTag?.innerXml ?? "";
  const [patternFillTag, remainingXml] = takeFirstXmlTagByName(innerXml, "patternFill");
  if (!patternFillTag) {
    return {
      definition: buildEmptyFillDefinition(),
      extraChildrenXml: /\S/.test(innerXml) ? innerXml : "",
    };
  }

  const patternAttributes = parseAttributes(patternFillTag.attributesSource);
  const patternInnerXml = patternFillTag.innerXml ?? "";
  const fgColorTag = findFirstXmlTag(patternInnerXml, "fgColor");
  const bgColorTag = findFirstXmlTag(patternInnerXml, "bgColor");

  return {
    definition: {
      patternType: findAttributeValue(patternAttributes, "patternType") ?? null,
      fgColor: parseFillColorDefinition(fgColorTag?.source ?? null),
      bgColor: parseFillColorDefinition(bgColorTag?.source ?? null),
    },
    extraChildrenXml: /\S/.test(remainingXml) ? remainingXml : "",
  };
}

function parseFont(fontXml: string): ParsedFont {
  const fontTag = findFirstXmlTag(fontXml, "font");
  const innerXml = fontTag?.innerXml ?? "";
  let remainingXml = innerXml;

  const [boldTag, remainingAfterBold] = takeFirstXmlTagByName(remainingXml, "b");
  remainingXml = remainingAfterBold;
  const [italicTag, remainingAfterItalic] = takeFirstXmlTagByName(remainingXml, "i");
  remainingXml = remainingAfterItalic;
  const [underlineTag, remainingAfterUnderline] = takeFirstXmlTagByName(remainingXml, "u");
  remainingXml = remainingAfterUnderline;
  const [strikeTag, remainingAfterStrike] = takeFirstXmlTagByName(remainingXml, "strike");
  remainingXml = remainingAfterStrike;
  const [outlineTag, remainingAfterOutline] = takeFirstXmlTagByName(remainingXml, "outline");
  remainingXml = remainingAfterOutline;
  const [shadowTag, remainingAfterShadow] = takeFirstXmlTagByName(remainingXml, "shadow");
  remainingXml = remainingAfterShadow;
  const [condenseTag, remainingAfterCondense] = takeFirstXmlTagByName(remainingXml, "condense");
  remainingXml = remainingAfterCondense;
  const [extendTag, remainingAfterExtend] = takeFirstXmlTagByName(remainingXml, "extend");
  remainingXml = remainingAfterExtend;
  const [colorTag, remainingAfterColor] = takeFirstXmlTagByName(remainingXml, "color");
  remainingXml = remainingAfterColor;
  const [sizeTag, remainingAfterSize] = takeFirstXmlTagByName(remainingXml, "sz");
  remainingXml = remainingAfterSize;
  const [nameTag, remainingAfterName] = takeFirstXmlTagByName(remainingXml, "name");
  remainingXml = remainingAfterName;
  const [familyTag, remainingAfterFamily] = takeFirstXmlTagByName(remainingXml, "family");
  remainingXml = remainingAfterFamily;
  const [charsetTag, remainingAfterCharset] = takeFirstXmlTagByName(remainingXml, "charset");
  remainingXml = remainingAfterCharset;
  const [schemeTag, remainingAfterScheme] = takeFirstXmlTagByName(remainingXml, "scheme");
  remainingXml = remainingAfterScheme;
  const [vertAlignTag, remainingAfterVertAlign] = takeFirstXmlTagByName(remainingXml, "vertAlign");
  remainingXml = remainingAfterVertAlign;

  return {
    definition: {
      bold: boldTag ? true : null,
      italic: italicTag ? true : null,
      underline: parseUnderlineValue(underlineTag?.source ?? null),
      strike: strikeTag ? true : null,
      outline: outlineTag ? true : null,
      shadow: shadowTag ? true : null,
      condense: condenseTag ? true : null,
      extend: extendTag ? true : null,
      size: parseTagValNumber(sizeTag?.source ?? null),
      name: parseTagValString(nameTag?.source ?? null),
      family: parseTagValNumber(familyTag?.source ?? null),
      charset: parseTagValNumber(charsetTag?.source ?? null),
      scheme: parseTagValString(schemeTag?.source ?? null),
      vertAlign: parseTagValString(vertAlignTag?.source ?? null),
      color: parseFontColorDefinition(colorTag?.source ?? null),
    },
    extraChildrenXml: /\S/.test(remainingXml) ? remainingXml : "",
  };
}

function parseCellStyle(attributesSource: string, innerXml: string): ParsedCellStyle {
  const attributes = parseAttributes(attributesSource);
  const [alignmentTag, extraChildrenXml] = takeFirstXmlTagByName(innerXml, "alignment");
  const alignmentAttributes = alignmentTag ? parseAttributes(alignmentTag.attributesSource) : null;

  return {
    alignmentAttributes,
    attributes,
    definition: {
      numFmtId: parseRequiredIntegerAttribute(attributes, "numFmtId", 0),
      fontId: parseRequiredIntegerAttribute(attributes, "fontId", 0),
      fillId: parseRequiredIntegerAttribute(attributes, "fillId", 0),
      borderId: parseRequiredIntegerAttribute(attributes, "borderId", 0),
      xfId: parseOptionalIntegerAttribute(attributes, "xfId"),
      quotePrefix: parseOptionalBooleanAttribute(attributes, "quotePrefix"),
      pivotButton: parseOptionalBooleanAttribute(attributes, "pivotButton"),
      applyNumberFormat: parseOptionalBooleanAttribute(attributes, "applyNumberFormat"),
      applyFont: parseOptionalBooleanAttribute(attributes, "applyFont"),
      applyFill: parseOptionalBooleanAttribute(attributes, "applyFill"),
      applyBorder: parseOptionalBooleanAttribute(attributes, "applyBorder"),
      applyAlignment: parseOptionalBooleanAttribute(attributes, "applyAlignment"),
      applyProtection: parseOptionalBooleanAttribute(attributes, "applyProtection"),
      alignment: alignmentAttributes ? parseAlignmentDefinition(alignmentAttributes) : null,
    },
    extraChildrenXml,
  };
}

function parseAlignmentDefinition(attributes: Array<[string, string]>): CellStyleAlignment {
  const alignment: CellStyleAlignment = {};

  assignStringAttribute(alignment, "horizontal", findAttributeValue(attributes, "horizontal"));
  assignStringAttribute(alignment, "vertical", findAttributeValue(attributes, "vertical"));
  assignNumberAttribute(alignment, "textRotation", findAttributeValue(attributes, "textRotation"));
  assignBooleanAttribute(alignment, "wrapText", findAttributeValue(attributes, "wrapText"));
  assignBooleanAttribute(alignment, "shrinkToFit", findAttributeValue(attributes, "shrinkToFit"));
  assignNumberAttribute(alignment, "indent", findAttributeValue(attributes, "indent"));
  assignNumberAttribute(alignment, "relativeIndent", findAttributeValue(attributes, "relativeIndent"));
  assignBooleanAttribute(alignment, "justifyLastLine", findAttributeValue(attributes, "justifyLastLine"));
  assignNumberAttribute(alignment, "readingOrder", findAttributeValue(attributes, "readingOrder"));

  return alignment;
}

function buildEmptyFillDefinition(): CellFillDefinition {
  return {
    patternType: null,
    fgColor: null,
    bgColor: null,
  };
}

function findAttributeValue(attributes: Array<[string, string]>, name: string): string | undefined {
  return attributes.find(([attributeName]) => attributeName === name)?.[1];
}

function takeFirstXmlTagByName(xml: string, tagName: string): [XmlTag | null, string] {
  const tag = findFirstXmlTag(xml, tagName);
  if (!tag) {
    return [null, xml];
  }

  return [tag, xml.slice(0, tag.start) + xml.slice(tag.end)];
}

function parseTagValNumber(tagXml: string | null): number | null {
  if (!tagXml) {
    return null;
  }
  const value = getXmlAttr(tagXml, "val");
  return value === undefined ? null : Number(value);
}

function parseTagValString(tagXml: string | null): string | null {
  if (!tagXml) {
    return null;
  }
  return getXmlAttr(tagXml, "val") ?? null;
}

function parseUnderlineValue(tagXml: string | null): string | null {
  if (!tagXml) {
    return null;
  }
  return getXmlAttr(tagXml, "val") ?? "single";
}

function parseFontColorDefinition(tagXml: string | null): CellFontColor | null {
  if (!tagXml) {
    return null;
  }

  const color: CellFontColor = {};
  const rgb = getXmlAttr(tagXml, "rgb");
  const theme = getXmlAttr(tagXml, "theme");
  const indexed = getXmlAttr(tagXml, "indexed");
  const auto = getXmlAttr(tagXml, "auto");
  const tint = getXmlAttr(tagXml, "tint");

  if (rgb !== undefined) {
    color.rgb = rgb;
  }
  if (theme !== undefined) {
    color.theme = Number(theme);
  }
  if (indexed !== undefined) {
    color.indexed = Number(indexed);
  }
  if (auto !== undefined) {
    color.auto = auto === "1" || auto === "true";
  }
  if (tint !== undefined) {
    color.tint = Number(tint);
  }

  return Object.keys(color).length === 0 ? null : color;
}

function parseFillColorDefinition(tagXml: string | null): CellFillColor | null {
  if (!tagXml) {
    return null;
  }

  const color: CellFillColor = {};
  const rgb = getXmlAttr(tagXml, "rgb");
  const theme = getXmlAttr(tagXml, "theme");
  const indexed = getXmlAttr(tagXml, "indexed");
  const auto = getXmlAttr(tagXml, "auto");
  const tint = getXmlAttr(tagXml, "tint");

  if (rgb !== undefined) {
    color.rgb = rgb;
  }
  if (theme !== undefined) {
    color.theme = Number(theme);
  }
  if (indexed !== undefined) {
    color.indexed = Number(indexed);
  }
  if (auto !== undefined) {
    color.auto = auto === "1" || auto === "true";
  }
  if (tint !== undefined) {
    color.tint = Number(tint);
  }

  return Object.keys(color).length === 0 ? null : color;
}

function parseBorderSideDefinition(tagXml: string | null): CellBorderSideDefinition | null {
  if (!tagXml) {
    return null;
  }

  const style = getXmlAttr(tagXml, "style") ?? null;
  const colorTag = findFirstXmlTag(tagXml, "color");
  const color = parseBorderColorDefinition(colorTag?.source ?? null);
  return {
    style,
    color,
  };
}

function parseBorderColorDefinition(tagXml: string | null): CellBorderColor | null {
  if (!tagXml) {
    return null;
  }

  const color: CellBorderColor = {};
  const rgb = getXmlAttr(tagXml, "rgb");
  const theme = getXmlAttr(tagXml, "theme");
  const indexed = getXmlAttr(tagXml, "indexed");
  const auto = getXmlAttr(tagXml, "auto");
  const tint = getXmlAttr(tagXml, "tint");

  if (rgb !== undefined) {
    color.rgb = rgb;
  }
  if (theme !== undefined) {
    color.theme = Number(theme);
  }
  if (indexed !== undefined) {
    color.indexed = Number(indexed);
  }
  if (auto !== undefined) {
    color.auto = auto === "1" || auto === "true";
  }
  if (tint !== undefined) {
    color.tint = Number(tint);
  }

  return Object.keys(color).length === 0 ? null : color;
}

function parseRequiredIntegerAttribute(attributes: Array<[string, string]>, name: string, fallback: number): number {
  const value = findAttributeValue(attributes, name);
  return value === undefined ? fallback : Number(value);
}

function parseOptionalIntegerAttribute(attributes: Array<[string, string]>, name: string): number | null {
  const value = findAttributeValue(attributes, name);
  return value === undefined ? null : Number(value);
}

function parseOptionalBooleanAttribute(attributes: Array<[string, string]>, name: string): boolean | null {
  const value = findAttributeValue(attributes, name);
  if (value === undefined) {
    return null;
  }

  return value === "1" || value === "true";
}

function assignStringAttribute(target: CellStyleAlignment, name: keyof CellStyleAlignment, value?: string): void {
  if (value !== undefined) {
    target[name] = value as never;
  }
}

function assignNumberAttribute(target: CellStyleAlignment, name: keyof CellStyleAlignment, value?: string): void {
  if (value !== undefined) {
    target[name] = Number(value) as never;
  }
}

function assignBooleanAttribute(target: CellStyleAlignment, name: keyof CellStyleAlignment, value?: string): void {
  if (value !== undefined) {
    target[name] = (value === "1" || value === "true") as never;
  }
}
