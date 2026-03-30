import { XlsxError } from "../errors.js";
import type {
  CellBorderPatch,
  CellFillPatch,
  CellFontPatch,
  CellStyleAlignmentPatch,
  CellStylePatch,
} from "../types.js";

export function assertStyleId(styleId: number | null): void {
  if (styleId !== null && (!Number.isInteger(styleId) || styleId < 0)) {
    throw new XlsxError(`Invalid style id: ${styleId}`);
  }
}

export function resolveCloneStylePatch(
  addressOrRowNumber: string | number,
  columnOrPatch: number | string | CellStylePatch | undefined,
  patch: CellStylePatch | undefined,
): CellStylePatch {
  return typeof addressOrRowNumber === "number" ? (patch ?? {}) : ((columnOrPatch as CellStylePatch | undefined) ?? {});
}

export function resolveSetAlignmentPatch(
  addressOrRowNumber: string | number,
  columnOrPatch: number | string | CellStyleAlignmentPatch | null,
  patch: CellStyleAlignmentPatch | null | undefined,
): CellStyleAlignmentPatch | null {
  if (typeof addressOrRowNumber === "number") {
    return patch === undefined ? {} : patch;
  }

  return columnOrPatch as CellStyleAlignmentPatch | null;
}

export function resolveSetFontPatch(
  addressOrRowNumber: string | number,
  columnOrPatch: number | string | CellFontPatch,
  patch: CellFontPatch | undefined,
): CellFontPatch {
  return typeof addressOrRowNumber === "number" ? (patch ?? {}) : (columnOrPatch as CellFontPatch);
}

export function resolveSetFillPatch(
  addressOrRowNumber: string | number,
  columnOrPatch: number | string | CellFillPatch,
  patch: CellFillPatch | undefined,
): CellFillPatch {
  return typeof addressOrRowNumber === "number" ? (patch ?? {}) : (columnOrPatch as CellFillPatch);
}

export function resolveSetBorderPatch(
  addressOrRowNumber: string | number,
  columnOrPatch: number | string | CellBorderPatch,
  patch: CellBorderPatch | undefined,
): CellBorderPatch {
  return typeof addressOrRowNumber === "number" ? (patch ?? {}) : (columnOrPatch as CellBorderPatch);
}

export function resolveSetStyleId(
  addressOrRowNumber: string | number,
  columnOrStyleId: number | string | null,
  styleId?: number | null,
): number | null {
  const nextStyleId =
    typeof addressOrRowNumber === "number" ? (styleId ?? null) : (columnOrStyleId as number | null);
  assertStyleId(nextStyleId);
  return nextStyleId;
}
