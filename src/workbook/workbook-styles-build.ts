import type {
  CellBorderColor,
  CellBorderColorPatch,
  CellBorderDefinition,
  CellBorderPatch,
  CellBorderSideDefinition,
  CellBorderSidePatch,
  CellFillColor,
  CellFillColorPatch,
  CellFillDefinition,
  CellFillPatch,
  CellFontColor,
  CellFontColorPatch,
  CellFontDefinition,
  CellFontPatch,
  CellStyleDefinition,
  CellStyleAlignmentPatch,
  CellStylePatch,
} from "../types.js";
import { serializeAttributes } from "../utils/xml.js";
import type { ParsedBorder, ParsedCellStyle, ParsedFill, ParsedFont } from "./workbook-styles-parse.js";

export function buildPatchedCellXfXml(sourceStyle: ParsedCellStyle, patch: CellStylePatch): string {
  const attributes = applyCellStylePatch(sourceStyle.attributes, patch);
  const alignmentAttributes = applyAlignmentPatch(sourceStyle.alignmentAttributes, patch.alignment);
  const alignmentXml = alignmentAttributes ? buildSelfClosingTag("alignment", alignmentAttributes) : "";
  const innerXml = alignmentXml + sourceStyle.extraChildrenXml;
  const serializedAttributes = serializeAttributes(attributes);

  if (innerXml.length === 0) {
    return `<xf${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
  }

  return `<xf${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</xf>`;
}

export function buildPatchedFontXml(sourceFont: ParsedFont, patch: CellFontPatch): string {
  const font = applyFontPatch(sourceFont.definition, patch);
  const childXml = buildFontChildXml(font) + sourceFont.extraChildrenXml;
  return childXml.length === 0 ? "<font/>" : `<font>${childXml}</font>`;
}

export function buildPatchedFillXml(sourceFill: ParsedFill, patch: CellFillPatch): string {
  const fill = applyFillPatch(sourceFill.definition, patch);
  const childXml = buildFillChildXml(fill) + sourceFill.extraChildrenXml;
  return childXml.length === 0 ? "<fill/>" : `<fill>${childXml}</fill>`;
}

export function buildPatchedBorderXml(sourceBorder: ParsedBorder, patch: CellBorderPatch): string {
  const border = applyBorderPatch(sourceBorder.definition, patch);
  const attributes = buildBorderAttributes(border);
  const serializedAttributes = serializeAttributes(attributes);
  const childXml = buildBorderChildXml(border) + sourceBorder.extraChildrenXml;
  return childXml.length === 0
    ? `<border${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`
    : `<border${serializedAttributes ? ` ${serializedAttributes}` : ""}>${childXml}</border>`;
}

export function cloneCellStyleDefinition(style: CellStyleDefinition | null): CellStyleDefinition | null {
  if (!style) {
    return null;
  }

  return {
    ...style,
    alignment: style.alignment ? { ...style.alignment } : null,
  };
}

export function cloneCellFontDefinition(font: CellFontDefinition | null): CellFontDefinition | null {
  if (!font) {
    return null;
  }

  return {
    ...font,
    color: font.color ? { ...font.color } : null,
  };
}

export function cloneCellFillDefinition(fill: CellFillDefinition | null): CellFillDefinition | null {
  if (!fill) {
    return null;
  }

  return {
    ...fill,
    fgColor: fill.fgColor ? { ...fill.fgColor } : null,
    bgColor: fill.bgColor ? { ...fill.bgColor } : null,
  };
}

export function cloneCellBorderDefinition(border: CellBorderDefinition | null): CellBorderDefinition | null {
  if (!border) {
    return null;
  }

  return {
    left: cloneCellBorderSideDefinition(border.left),
    right: cloneCellBorderSideDefinition(border.right),
    top: cloneCellBorderSideDefinition(border.top),
    bottom: cloneCellBorderSideDefinition(border.bottom),
    diagonal: cloneCellBorderSideDefinition(border.diagonal),
    vertical: cloneCellBorderSideDefinition(border.vertical),
    horizontal: cloneCellBorderSideDefinition(border.horizontal),
    diagonalUp: border.diagonalUp,
    diagonalDown: border.diagonalDown,
    outline: border.outline,
  };
}

export function getNextCustomNumberFormatId(numberFormats: Map<number, string>): number {
  let nextNumFmtId = 164;

  for (const numFmtId of numberFormats.keys()) {
    nextNumFmtId = Math.max(nextNumFmtId, numFmtId + 1);
  }

  return nextNumFmtId;
}

function applyCellStylePatch(attributes: Array<[string, string]>, patch: CellStylePatch): Array<[string, string]> {
  let nextAttributes = [...attributes];

  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "numFmtId", patch.numFmtId);
  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "fontId", patch.fontId);
  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "fillId", patch.fillId);
  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "borderId", patch.borderId);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "xfId", patch.xfId);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "quotePrefix", patch.quotePrefix);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "pivotButton", patch.pivotButton);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyNumberFormat", patch.applyNumberFormat);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyFont", patch.applyFont);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyFill", patch.applyFill);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyBorder", patch.applyBorder);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyAlignment", patch.applyAlignment);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyProtection", patch.applyProtection);

  return nextAttributes;
}

function applyAlignmentPatch(
  attributes: Array<[string, string]> | null,
  patch: CellStyleAlignmentPatch | null | undefined,
): Array<[string, string]> | null {
  if (patch === undefined) {
    return attributes ? [...attributes] : null;
  }

  if (patch === null) {
    return null;
  }

  let nextAttributes = attributes ? [...attributes] : [];
  nextAttributes = applyOptionalStringPatch(nextAttributes, "horizontal", patch.horizontal);
  nextAttributes = applyOptionalStringPatch(nextAttributes, "vertical", patch.vertical);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "textRotation", patch.textRotation);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "wrapText", patch.wrapText);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "shrinkToFit", patch.shrinkToFit);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "indent", patch.indent);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "relativeIndent", patch.relativeIndent);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "justifyLastLine", patch.justifyLastLine);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "readingOrder", patch.readingOrder);

  return nextAttributes.length === 0 ? null : nextAttributes;
}

function buildSelfClosingTag(tagName: string, attributes: Array<[string, string]>): string {
  const serializedAttributes = serializeAttributes(attributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
}

function buildFontChildXml(font: CellFontDefinition): string {
  const parts: string[] = [];

  if (font.bold) {
    parts.push("<b/>");
  }
  if (font.italic) {
    parts.push("<i/>");
  }
  if (font.underline !== null) {
    parts.push(font.underline === "single" ? "<u/>" : buildSelfClosingTag("u", [["val", font.underline]]));
  }
  if (font.strike) {
    parts.push("<strike/>");
  }
  if (font.outline) {
    parts.push("<outline/>");
  }
  if (font.shadow) {
    parts.push("<shadow/>");
  }
  if (font.condense) {
    parts.push("<condense/>");
  }
  if (font.extend) {
    parts.push("<extend/>");
  }
  if (font.color) {
    parts.push(buildSelfClosingTag("color", buildFontColorAttributes(font.color)));
  }
  if (font.size !== null) {
    parts.push(buildSelfClosingTag("sz", [["val", String(font.size)]]));
  }
  if (font.name !== null) {
    parts.push(buildSelfClosingTag("name", [["val", font.name]]));
  }
  if (font.family !== null) {
    parts.push(buildSelfClosingTag("family", [["val", String(font.family)]]));
  }
  if (font.charset !== null) {
    parts.push(buildSelfClosingTag("charset", [["val", String(font.charset)]]));
  }
  if (font.scheme !== null) {
    parts.push(buildSelfClosingTag("scheme", [["val", font.scheme]]));
  }
  if (font.vertAlign !== null) {
    parts.push(buildSelfClosingTag("vertAlign", [["val", font.vertAlign]]));
  }

  return parts.join("");
}

function buildFillChildXml(fill: CellFillDefinition): string {
  if (fill.patternType === null && fill.fgColor === null && fill.bgColor === null) {
    return "";
  }

  const attributes = fill.patternType === null ? [] : ([["patternType", fill.patternType]] as Array<[string, string]>);
  const colorXml =
    (fill.fgColor ? buildSelfClosingTag("fgColor", buildFillColorAttributes(fill.fgColor)) : "") +
    (fill.bgColor ? buildSelfClosingTag("bgColor", buildFillColorAttributes(fill.bgColor)) : "");

  if (colorXml.length === 0) {
    return buildSelfClosingTag("patternFill", attributes);
  }

  const serializedAttributes = serializeAttributes(attributes);
  return `<patternFill${serializedAttributes ? ` ${serializedAttributes}` : ""}>${colorXml}</patternFill>`;
}

function buildBorderChildXml(border: CellBorderDefinition): string {
  return [
    buildBorderSideXml("left", border.left),
    buildBorderSideXml("right", border.right),
    buildBorderSideXml("top", border.top),
    buildBorderSideXml("bottom", border.bottom),
    buildBorderSideXml("diagonal", border.diagonal),
    buildBorderSideXml("vertical", border.vertical),
    buildBorderSideXml("horizontal", border.horizontal),
  ].join("");
}

function buildBorderAttributes(border: CellBorderDefinition): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (border.diagonalUp !== null) {
    attributes.push(["diagonalUp", border.diagonalUp ? "1" : "0"]);
  }
  if (border.diagonalDown !== null) {
    attributes.push(["diagonalDown", border.diagonalDown ? "1" : "0"]);
  }
  if (border.outline !== null) {
    attributes.push(["outline", border.outline ? "1" : "0"]);
  }
  return attributes;
}

function buildBorderSideXml(tagName: string, side: CellBorderSideDefinition | null): string {
  if (side === null) {
    return "";
  }

  const attributes: Array<[string, string]> = [];
  if (side.style !== null) {
    attributes.push(["style", side.style]);
  }

  const colorXml = side.color ? buildSelfClosingTag("color", buildBorderColorAttributes(side.color)) : "";
  if (colorXml.length === 0) {
    return buildSelfClosingTag(tagName, attributes);
  }

  const serializedAttributes = serializeAttributes(attributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${colorXml}</${tagName}>`;
}

function buildEmptyFontDefinition(): CellFontDefinition {
  return {
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: null,
    name: null,
    family: null,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: null,
  };
}

function buildEmptyBorderDefinition(): CellBorderDefinition {
  return {
    left: null,
    right: null,
    top: null,
    bottom: null,
    diagonal: null,
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  };
}

function applyFontPatch(sourceFont: CellFontDefinition, patch: CellFontPatch): CellFontDefinition {
  return {
    bold: patch.bold === undefined ? sourceFont.bold : patch.bold,
    italic: patch.italic === undefined ? sourceFont.italic : patch.italic,
    underline: patch.underline === undefined ? sourceFont.underline : patch.underline,
    strike: patch.strike === undefined ? sourceFont.strike : patch.strike,
    outline: patch.outline === undefined ? sourceFont.outline : patch.outline,
    shadow: patch.shadow === undefined ? sourceFont.shadow : patch.shadow,
    condense: patch.condense === undefined ? sourceFont.condense : patch.condense,
    extend: patch.extend === undefined ? sourceFont.extend : patch.extend,
    size: patch.size === undefined ? sourceFont.size : patch.size,
    name: patch.name === undefined ? sourceFont.name : patch.name,
    family: patch.family === undefined ? sourceFont.family : patch.family,
    charset: patch.charset === undefined ? sourceFont.charset : patch.charset,
    scheme: patch.scheme === undefined ? sourceFont.scheme : patch.scheme,
    vertAlign: patch.vertAlign === undefined ? sourceFont.vertAlign : patch.vertAlign,
    color: applyFontColorPatch(sourceFont.color, patch.color),
  };
}

function applyFillPatch(sourceFill: CellFillDefinition, patch: CellFillPatch): CellFillDefinition {
  return {
    patternType: patch.patternType === undefined ? sourceFill.patternType : patch.patternType,
    fgColor: applyFillColorPatch(sourceFill.fgColor, patch.fgColor),
    bgColor: applyFillColorPatch(sourceFill.bgColor, patch.bgColor),
  };
}

function applyBorderPatch(sourceBorder: CellBorderDefinition, patch: CellBorderPatch): CellBorderDefinition {
  return {
    left: applyBorderSidePatch(sourceBorder.left, patch.left),
    right: applyBorderSidePatch(sourceBorder.right, patch.right),
    top: applyBorderSidePatch(sourceBorder.top, patch.top),
    bottom: applyBorderSidePatch(sourceBorder.bottom, patch.bottom),
    diagonal: applyBorderSidePatch(sourceBorder.diagonal, patch.diagonal),
    vertical: applyBorderSidePatch(sourceBorder.vertical, patch.vertical),
    horizontal: applyBorderSidePatch(sourceBorder.horizontal, patch.horizontal),
    diagonalUp: patch.diagonalUp === undefined ? sourceBorder.diagonalUp : patch.diagonalUp,
    diagonalDown: patch.diagonalDown === undefined ? sourceBorder.diagonalDown : patch.diagonalDown,
    outline: patch.outline === undefined ? sourceBorder.outline : patch.outline,
  };
}

function applyFontColorPatch(
  sourceColor: CellFontColor | null,
  patch: CellFontColorPatch | null | undefined,
): CellFontColor | null {
  if (patch === undefined) {
    return sourceColor ? { ...sourceColor } : null;
  }
  if (patch === null) {
    return null;
  }

  const nextColor: CellFontColor = sourceColor ? { ...sourceColor } : {};
  updateOptionalObjectProperty(nextColor, "rgb", patch.rgb);
  updateOptionalObjectProperty(nextColor, "theme", patch.theme);
  updateOptionalObjectProperty(nextColor, "indexed", patch.indexed);
  updateOptionalObjectProperty(nextColor, "auto", patch.auto);
  updateOptionalObjectProperty(nextColor, "tint", patch.tint);

  return Object.keys(nextColor).length === 0 ? null : nextColor;
}

function buildFontColorAttributes(color: CellFontColor): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (color.rgb !== undefined) {
    attributes.push(["rgb", color.rgb]);
  }
  if (color.theme !== undefined) {
    attributes.push(["theme", String(color.theme)]);
  }
  if (color.indexed !== undefined) {
    attributes.push(["indexed", String(color.indexed)]);
  }
  if (color.auto !== undefined) {
    attributes.push(["auto", color.auto ? "1" : "0"]);
  }
  if (color.tint !== undefined) {
    attributes.push(["tint", String(color.tint)]);
  }
  return attributes;
}

function applyFillColorPatch(
  sourceColor: CellFillColor | null,
  patch: CellFillColorPatch | null | undefined,
): CellFillColor | null {
  if (patch === undefined) {
    return sourceColor ? { ...sourceColor } : null;
  }
  if (patch === null) {
    return null;
  }

  const nextColor: CellFillColor = sourceColor ? { ...sourceColor } : {};
  updateOptionalObjectProperty(nextColor, "rgb", patch.rgb);
  updateOptionalObjectProperty(nextColor, "theme", patch.theme);
  updateOptionalObjectProperty(nextColor, "indexed", patch.indexed);
  updateOptionalObjectProperty(nextColor, "auto", patch.auto);
  updateOptionalObjectProperty(nextColor, "tint", patch.tint);

  return Object.keys(nextColor).length === 0 ? null : colorless(nextColor);
}

function applyBorderSidePatch(
  sourceSide: CellBorderSideDefinition | null,
  patch: CellBorderSidePatch | null | undefined,
): CellBorderSideDefinition | null {
  if (patch === undefined) {
    return cloneCellBorderSideDefinition(sourceSide);
  }
  if (patch === null) {
    return null;
  }

  return {
    style: patch.style === undefined ? (sourceSide?.style ?? null) : patch.style,
    color: applyBorderColorPatch(sourceSide?.color ?? null, patch.color),
  };
}

function applyBorderColorPatch(
  sourceColor: CellBorderColor | null,
  patch: CellBorderColorPatch | null | undefined,
): CellBorderColor | null {
  if (patch === undefined) {
    return sourceColor ? { ...sourceColor } : null;
  }
  if (patch === null) {
    return null;
  }

  const nextColor: CellBorderColor = sourceColor ? { ...sourceColor } : {};
  updateOptionalObjectProperty(nextColor, "rgb", patch.rgb);
  updateOptionalObjectProperty(nextColor, "theme", patch.theme);
  updateOptionalObjectProperty(nextColor, "indexed", patch.indexed);
  updateOptionalObjectProperty(nextColor, "auto", patch.auto);
  updateOptionalObjectProperty(nextColor, "tint", patch.tint);

  return Object.keys(nextColor).length === 0 ? null : colorless(nextColor);
}

function buildFillColorAttributes(color: CellFillColor): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (color.rgb !== undefined) {
    attributes.push(["rgb", color.rgb]);
  }
  if (color.theme !== undefined) {
    attributes.push(["theme", String(color.theme)]);
  }
  if (color.indexed !== undefined) {
    attributes.push(["indexed", String(color.indexed)]);
  }
  if (color.auto !== undefined) {
    attributes.push(["auto", color.auto ? "1" : "0"]);
  }
  if (color.tint !== undefined) {
    attributes.push(["tint", String(color.tint)]);
  }
  return attributes;
}

function buildBorderColorAttributes(color: CellBorderColor): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (color.rgb !== undefined) {
    attributes.push(["rgb", color.rgb]);
  }
  if (color.theme !== undefined) {
    attributes.push(["theme", String(color.theme)]);
  }
  if (color.indexed !== undefined) {
    attributes.push(["indexed", String(color.indexed)]);
  }
  if (color.auto !== undefined) {
    attributes.push(["auto", color.auto ? "1" : "0"]);
  }
  if (color.tint !== undefined) {
    attributes.push(["tint", String(color.tint)]);
  }
  return attributes;
}

function updateOptionalObjectProperty<T extends object, K extends keyof T>(
  target: T,
  key: K,
  value: T[K] | null | undefined,
): void {
  if (value === undefined) {
    return;
  }

  if (value === null) {
    delete target[key];
    return;
  }

  target[key] = value;
}

function cloneCellBorderSideDefinition(side: CellBorderSideDefinition | null): CellBorderSideDefinition | null {
  if (!side) {
    return null;
  }

  return {
    style: side.style,
    color: side.color ? { ...side.color } : null,
  };
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

function applyRequiredIntegerPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: number | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, String(value));
}

function applyOptionalIntegerPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: number | null | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, value === null ? null : String(value));
}

function applyOptionalBooleanPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: boolean | null | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, value === null ? null : value ? "1" : "0");
}

function applyOptionalStringPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: string | null | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, value);
}

function colorless<T extends object>(value: T): T {
  return value;
}
