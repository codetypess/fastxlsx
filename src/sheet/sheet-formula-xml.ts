import type { CellValue } from "../types.js";
import { XlsxError } from "../errors.js";
import { decodeXmlText, escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";

export function parseFormulaTagInfo(cellXml: string): { attributesSource: string; formula: string | null } {
  const { innerXml } = splitCellXml(cellXml);
  const formulaStart = innerXml.indexOf("<f");
  if (formulaStart === -1) {
    return { attributesSource: "", formula: null };
  }

  const formulaOpenTagEnd = innerXml.indexOf(">", formulaStart + 2);
  if (formulaOpenTagEnd === -1) {
    throw new XlsxError("Formula cell XML is missing a closing <f> tag opener");
  }

  const attributesSource = innerXml.slice(formulaStart + 2, formulaOpenTagEnd).trim();
  if (isSelfClosingTagSource(innerXml.slice(formulaStart + 2, formulaOpenTagEnd))) {
    return { attributesSource, formula: null };
  }

  const formulaCloseStart = innerXml.indexOf("</f>", formulaOpenTagEnd + 1);
  if (formulaCloseStart === -1) {
    throw new XlsxError("Formula cell XML is missing </f>");
  }

  return {
    attributesSource,
    formula: decodeXmlText(innerXml.slice(formulaOpenTagEnd + 1, formulaCloseStart)),
  };
}

export function buildRecalculatedFormulaCellXml(
  address: string,
  cellXml: string,
  value: CellValue,
  errorText: string | null,
): string {
  const { attributesSource, innerXml } = splitCellXml(cellXml);
  const attributes = parseAttributes(attributesSource);
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "t");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (errorText !== null) {
    nextAttributes.push(["t", "e"]);
  } else if (typeof value === "string") {
    nextAttributes.push(["t", "str"]);
  } else if (typeof value === "boolean") {
    nextAttributes.push(["t", "b"]);
  }

  nextAttributes.push(...preserved);

  const formulaStart = innerXml.indexOf("<f");
  if (formulaStart === -1) {
    throw new XlsxError("Formula cell XML is missing <f>");
  }

  const formulaOpenTagEnd = innerXml.indexOf(">", formulaStart + 2);
  if (formulaOpenTagEnd === -1) {
    throw new XlsxError("Formula cell XML is missing a closing <f> tag opener");
  }

  if (isSelfClosingTagSource(innerXml.slice(formulaStart + 2, formulaOpenTagEnd))) {
    throw new XlsxError("Self-closing formula tags are not supported");
  }

  const formulaCloseEnd = innerXml.indexOf("</f>", formulaOpenTagEnd + 1);
  if (formulaCloseEnd === -1) {
    throw new XlsxError("Formula cell XML is missing </f>");
  }

  const formulaTagEnd = formulaCloseEnd + 4;
  const innerWithoutValue = innerXml.replace(/<v(?:\s[^>]*)?\/>|<v(?:\s[^>]*)?>[\s\S]*?<\/v>/g, "");
  const valueXml = buildFormulaValueXml(value, errorText);
  const nextInnerXml =
    valueXml.length === 0
      ? innerWithoutValue
      : innerWithoutValue.slice(0, formulaTagEnd) + valueXml + innerWithoutValue.slice(formulaTagEnd);

  return `<c ${serializeAttributes(nextAttributes)}>${nextInnerXml}</c>`;
}

function splitCellXml(cellXml: string): { attributesSource: string; innerXml: string } {
  const openTagEnd = cellXml.indexOf(">");
  if (openTagEnd === -1) {
    throw new XlsxError("Cell XML is missing opening tag");
  }

  const closeTagStart = cellXml.lastIndexOf("</c>");
  if (closeTagStart === -1) {
    throw new XlsxError("Cell XML is missing closing tag");
  }

  return {
    attributesSource: cellXml.slice(2, openTagEnd).trim(),
    innerXml: cellXml.slice(openTagEnd + 1, closeTagStart),
  };
}

function buildFormulaValueXml(value: CellValue, errorText: string | null): string {
  if (errorText !== null) {
    return `<v>${escapeXmlText(errorText)}</v>`;
  }

  if (value === null) {
    return "";
  }

  if (typeof value === "string") {
    return `<v>${escapeXmlText(value)}</v>`;
  }

  if (typeof value === "boolean") {
    return `<v>${value ? "1" : "0"}</v>`;
  }

  return `<v>${String(value)}</v>`;
}

function isSelfClosingTagSource(source: string): boolean {
  for (let index = source.length - 1; index >= 0; index -= 1) {
    const character = source[index];
    if (!character || /\s/.test(character)) {
      continue;
    }

    return character === "/";
  }

  return false;
}
