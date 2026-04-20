import type { Sheet } from "./sheet.js";
import type {
  CellStyleAlignment,
  CellStyleAlignmentPatch,
  CellBorderDefinition,
  CellBorderPatch,
  CellFillDefinition,
  CellFillPatch,
  CellFontDefinition,
  CellFontPatch,
  CellError,
  CellNumberFormatDefinition,
  CellSnapshot,
  CellStyleDefinition,
  CellStylePatch,
  CellType,
  CellValue,
  SetFormulaOptions,
} from "./types.js";

/**
 * Address-based cell handle returned by `Sheet#cell()`.
 *
 * The handle stays bound to the parent worksheet and always resolves the
 * latest cell state on demand.
 */
export class Cell {
  /**
   * Normalized A1-style address for this handle.
   */
  readonly address: string;

  private cachedRevision = -1;
  private cachedSnapshot?: CellSnapshot;
  private readonly sheet: Sheet;

  /**
   * Creates a cell handle for a normalized worksheet address.
   *
   * Most callers should use `Sheet#cell()` instead of instantiating `Cell`
   * directly.
   */
  constructor(sheet: Sheet, address: string) {
    this.sheet = sheet;
    this.address = address;
  }

  /**
   * Whether the cell node currently exists in the worksheet XML.
   */
  get exists(): boolean {
    return this.getSnapshot().exists;
  }

  /**
   * Formula text when the cell is a formula cell.
   */
  get formula(): string | null {
    return this.getSnapshot().formula;
  }

  /**
   * Structured Excel error metadata for cells with `t="e"` cached values.
   */
  get error(): CellError | null {
    return this.getSnapshot().error;
  }

  /**
   * Raw OOXML cell type attribute, if present.
   */
  get rawType(): string | null {
    return this.getSnapshot().rawType;
  }

  /**
   * Raw OOXML style id assigned to the cell.
   */
  get styleId(): number | null {
    return this.getSnapshot().styleId;
  }

  /**
   * Resolved cell style definition.
   */
  get style(): CellStyleDefinition | null {
    return this.sheet.getStyle(this.address);
  }

  /**
   * Resolved alignment definition from the current style.
   */
  get alignment(): CellStyleAlignment | null {
    return this.sheet.getAlignment(this.address);
  }

  /**
   * Resolved font definition from the current style.
   */
  get font(): CellFontDefinition | null {
    return this.sheet.getFont(this.address);
  }

  /**
   * Resolved fill definition from the current style.
   */
  get fill(): CellFillDefinition | null {
    return this.sheet.getFill(this.address);
  }

  /**
   * Convenience accessor for solid background color.
   */
  get backgroundColor(): string | null {
    return this.sheet.getBackgroundColor(this.address);
  }

  /**
   * Resolved border definition from the current style.
   */
  get border(): CellBorderDefinition | null {
    return this.sheet.getBorder(this.address);
  }

  /**
   * Resolved number format definition from the current style.
   */
  get numberFormat(): CellNumberFormatDefinition | null {
    return this.sheet.getNumberFormat(this.address);
  }

  /**
   * Normalized high-level cell value type.
   */
  get type(): CellType {
    return this.getSnapshot().type;
  }

  /**
   * Decoded cell value.
   */
  get value(): CellValue {
    return this.getSnapshot().value;
  }

  /**
   * Best-effort display text for the current cell value.
   */
  get text(): string | null {
    return this.sheet.getDisplayValue(this.address);
  }

  /**
   * Manually recalculates this formula cell.
   */
  recalculate(): CellSnapshot {
    return this.sheet.recalculateCell(this.address);
  }

  /**
   * Writes a formula to this cell.
   */
  setFormula(formula: string, options: SetFormulaOptions = {}): void {
    this.sheet.setFormula(this.address, formula, options);
  }

  /**
   * Writes a value to this cell.
   */
  setValue(value: CellValue): void {
    this.sheet.setCell(this.address, value);
  }

  /**
   * Assigns a raw style id to this cell.
   */
  setStyleId(styleId: number | null): void {
    this.sheet.setStyleId(this.address, styleId);
  }

  /**
   * Clones the current style with a patch and applies the new style id.
   */
  setStyle(patch: CellStylePatch): number {
    return this.sheet.setStyle(this.address, patch);
  }

  /**
   * Updates only the alignment portion of the current style.
   */
  setAlignment(patch: CellStyleAlignmentPatch | null): number {
    return this.sheet.setAlignment(this.address, patch);
  }

  /**
   * Updates only the font portion of the current style.
   */
  setFont(patch: CellFontPatch): number {
    return this.sheet.setFont(this.address, patch);
  }

  /**
   * Updates only the fill portion of the current style.
   */
  setFill(patch: CellFillPatch): number {
    return this.sheet.setFill(this.address, patch);
  }

  /**
   * Convenience helper for setting a solid fill color.
   */
  setBackgroundColor(color: string | null): number {
    return this.sheet.setBackgroundColor(this.address, color);
  }

  /**
   * Updates only the border portion of the current style.
   */
  setBorder(patch: CellBorderPatch): number {
    return this.sheet.setBorder(this.address, patch);
  }

  /**
   * Updates only the number format portion of the current style.
   */
  setNumberFormat(formatCode: string): number {
    return this.sheet.setNumberFormat(this.address, formatCode);
  }

  /**
   * Clones the current style and returns the new style id.
   */
  cloneStyle(patch: CellStylePatch = {}): number {
    return this.sheet.cloneStyle(this.address, patch);
  }

  private getSnapshot(): CellSnapshot {
    const revision = this.sheet.getRevision();
    if (this.cachedSnapshot && this.cachedRevision === revision) {
      return this.cachedSnapshot;
    }

    this.cachedSnapshot = this.sheet.readCellSnapshot(this.address);
    this.cachedRevision = revision;
    return this.cachedSnapshot;
  }
}
