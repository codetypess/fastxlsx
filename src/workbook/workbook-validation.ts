import { XlsxError } from "../errors.js";
import type {
  CellBorderColorPatch,
  CellBorderPatch,
  CellBorderSidePatch,
  CellFillColorPatch,
  CellFillPatch,
  CellFontColorPatch,
  CellFontPatch,
  CellStyleAlignmentPatch,
  CellStylePatch,
  SheetVisibility,
} from "../types.js";

export function assertSheetName(sheetName: string): void {
  if (sheetName.length === 0 || sheetName.length > 31 || /[\\/*?:[\]]/.test(sheetName)) {
    throw new XlsxError(`Invalid sheet name: ${sheetName}`);
  }
}

export function assertStyleId(styleId: number): void {
  if (!Number.isInteger(styleId) || styleId < 0) {
    throw new XlsxError(`Invalid style id: ${styleId}`);
  }
}

export function assertCellStylePatch(patch: CellStylePatch): void {
  assertOptionalNonNegativeInteger(patch.numFmtId, "numFmtId");
  assertOptionalNonNegativeInteger(patch.fontId, "fontId");
  assertOptionalNonNegativeInteger(patch.fillId, "fillId");
  assertOptionalNonNegativeInteger(patch.borderId, "borderId");
  assertOptionalNullableNonNegativeInteger(patch.xfId, "xfId");
  assertOptionalNullableBoolean(patch.quotePrefix, "quotePrefix");
  assertOptionalNullableBoolean(patch.pivotButton, "pivotButton");
  assertOptionalNullableBoolean(patch.applyNumberFormat, "applyNumberFormat");
  assertOptionalNullableBoolean(patch.applyFont, "applyFont");
  assertOptionalNullableBoolean(patch.applyFill, "applyFill");
  assertOptionalNullableBoolean(patch.applyBorder, "applyBorder");
  assertOptionalNullableBoolean(patch.applyAlignment, "applyAlignment");
  assertOptionalNullableBoolean(patch.applyProtection, "applyProtection");

  if (patch.alignment !== undefined && patch.alignment !== null) {
    assertCellStyleAlignmentPatch(patch.alignment);
  }
}

export function assertCellFontPatch(patch: CellFontPatch): void {
  assertOptionalNullableBoolean(patch.bold, "bold");
  assertOptionalNullableBoolean(patch.italic, "italic");
  assertOptionalNullableString(patch.underline, "underline");
  assertOptionalNullableBoolean(patch.strike, "strike");
  assertOptionalNullableBoolean(patch.outline, "outline");
  assertOptionalNullableBoolean(patch.shadow, "shadow");
  assertOptionalNullableBoolean(patch.condense, "condense");
  assertOptionalNullableBoolean(patch.extend, "extend");
  assertOptionalNullableFiniteNumber(patch.size, "size");
  assertOptionalNullableString(patch.name, "name");
  assertOptionalNullableNonNegativeInteger(patch.family, "family");
  assertOptionalNullableNonNegativeInteger(patch.charset, "charset");
  assertOptionalNullableString(patch.scheme, "scheme");
  assertOptionalNullableString(patch.vertAlign, "vertAlign");

  if (patch.color !== undefined && patch.color !== null) {
    assertCellFontColorPatch(patch.color);
  }
}

export function assertCellFillPatch(patch: CellFillPatch): void {
  assertOptionalNullableString(patch.patternType, "patternType");

  if (patch.fgColor !== undefined && patch.fgColor !== null) {
    assertCellFillColorPatch(patch.fgColor, "fgColor");
  }

  if (patch.bgColor !== undefined && patch.bgColor !== null) {
    assertCellFillColorPatch(patch.bgColor, "bgColor");
  }
}

export function assertCellBorderPatch(patch: CellBorderPatch): void {
  assertOptionalNullableBoolean(patch.diagonalUp, "diagonalUp");
  assertOptionalNullableBoolean(patch.diagonalDown, "diagonalDown");
  assertOptionalNullableBoolean(patch.outline, "outline");

  assertCellBorderSidePatch(patch.left, "left");
  assertCellBorderSidePatch(patch.right, "right");
  assertCellBorderSidePatch(patch.top, "top");
  assertCellBorderSidePatch(patch.bottom, "bottom");
  assertCellBorderSidePatch(patch.diagonal, "diagonal");
  assertCellBorderSidePatch(patch.vertical, "vertical");
  assertCellBorderSidePatch(patch.horizontal, "horizontal");
}

export function assertFormatCode(formatCode: string): void {
  if (formatCode.length === 0) {
    throw new XlsxError("Invalid format code: empty");
  }
}

export function assertDefinedName(name: string): void {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) {
    throw new XlsxError(`Invalid defined name: ${name}`);
  }
}

export function assertSheetVisibility(visibility: string): asserts visibility is SheetVisibility {
  if (visibility !== "visible" && visibility !== "hidden" && visibility !== "veryHidden") {
    throw new XlsxError(`Invalid sheet visibility: ${visibility}`);
  }
}

export function assertSheetIndex(sheetIndex: number, sheetCount: number): void {
  if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= sheetCount) {
    throw new XlsxError(`Invalid sheet index: ${sheetIndex}`);
  }
}

function assertCellFontColorPatch(patch: CellFontColorPatch): void {
  assertOptionalNullableString(patch.rgb, "color.rgb");
  assertOptionalNullableNonNegativeInteger(patch.theme, "color.theme");
  assertOptionalNullableNonNegativeInteger(patch.indexed, "color.indexed");
  assertOptionalNullableBoolean(patch.auto, "color.auto");
  assertOptionalNullableFiniteNumber(patch.tint, "color.tint");
}

function assertCellFillColorPatch(patch: CellFillColorPatch, name: string): void {
  assertOptionalNullableString(patch.rgb, `${name}.rgb`);
  assertOptionalNullableNonNegativeInteger(patch.theme, `${name}.theme`);
  assertOptionalNullableNonNegativeInteger(patch.indexed, `${name}.indexed`);
  assertOptionalNullableBoolean(patch.auto, `${name}.auto`);
  assertOptionalNullableFiniteNumber(patch.tint, `${name}.tint`);
}

function assertCellBorderSidePatch(patch: CellBorderSidePatch | null | undefined, name: string): void {
  if (patch === undefined || patch === null) {
    return;
  }

  assertOptionalNullableString(patch.style, `${name}.style`);
  if (patch.color !== undefined && patch.color !== null) {
    assertCellBorderColorPatch(patch.color, `${name}.color`);
  }
}

function assertCellBorderColorPatch(patch: CellBorderColorPatch, name: string): void {
  assertOptionalNullableString(patch.rgb, `${name}.rgb`);
  assertOptionalNullableNonNegativeInteger(patch.theme, `${name}.theme`);
  assertOptionalNullableNonNegativeInteger(patch.indexed, `${name}.indexed`);
  assertOptionalNullableBoolean(patch.auto, `${name}.auto`);
  assertOptionalNullableFiniteNumber(patch.tint, `${name}.tint`);
}

function assertCellStyleAlignmentPatch(patch: CellStyleAlignmentPatch): void {
  assertOptionalNullableString(patch.horizontal, "alignment.horizontal");
  assertOptionalNullableString(patch.vertical, "alignment.vertical");
  assertOptionalNullableNonNegativeInteger(patch.textRotation, "alignment.textRotation");
  assertOptionalNullableBoolean(patch.wrapText, "alignment.wrapText");
  assertOptionalNullableBoolean(patch.shrinkToFit, "alignment.shrinkToFit");
  assertOptionalNullableNonNegativeInteger(patch.indent, "alignment.indent");
  assertOptionalNullableNonNegativeInteger(patch.relativeIndent, "alignment.relativeIndent");
  assertOptionalNullableBoolean(patch.justifyLastLine, "alignment.justifyLastLine");
  assertOptionalNullableNonNegativeInteger(patch.readingOrder, "alignment.readingOrder");
}

function assertOptionalNonNegativeInteger(value: number | undefined, name: string): void {
  if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
    throw new XlsxError(`Invalid ${name}: ${value}`);
  }
}

function assertOptionalNullableNonNegativeInteger(value: number | null | undefined, name: string): void {
  if (value !== undefined && value !== null && (!Number.isInteger(value) || value < 0)) {
    throw new XlsxError(`Invalid ${name}: ${value}`);
  }
}

function assertOptionalNullableFiniteNumber(value: number | null | undefined, name: string): void {
  if (value !== undefined && value !== null && !Number.isFinite(value)) {
    throw new XlsxError(`Invalid ${name}: ${value}`);
  }
}

function assertOptionalNullableBoolean(value: boolean | null | undefined, name: string): void {
  if (value !== undefined && value !== null && typeof value !== "boolean") {
    throw new XlsxError(`Invalid ${name}: ${String(value)}`);
  }
}

function assertOptionalNullableString(value: string | null | undefined, name: string): void {
  if (value !== undefined && value !== null && typeof value !== "string") {
    throw new XlsxError(`Invalid ${name}: ${String(value)}`);
  }
}
