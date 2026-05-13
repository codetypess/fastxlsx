import { XlsxError } from "../errors.js";

export interface ValueWindowCellScanResult {
    addressSource: string;
    columnNumber: number;
    innerEnd: number;
    innerStart: number;
    logical: boolean;
    nextCursor: number;
    rawType: string | null;
}

export function cleanTagAttributesSource(source: string): string {
    let end = source.length;

    while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
        end -= 1;
    }

    if (end > 0 && source.charCodeAt(end - 1) === 47) {
        end -= 1;
    }

    while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
        end -= 1;
    }

    let start = 0;
    while (start < end && isXmlWhitespaceCode(source.charCodeAt(start))) {
        start += 1;
    }

    return source.slice(start, end);
}

export function hasLogicalValueWindowCellContent(
    xml: string,
    start: number,
    end: number,
    rawType: string | null
): boolean {
    return (
        hasCellChildTagFast(xml, start, end, "f") ||
        hasCellChildTagFast(xml, start, end, "v") ||
        (rawType === "inlineStr" && hasCellChildTagFast(xml, start, end, "is"))
    );
}

export function isSelfClosingTagSource(source: string): boolean {
    let index = source.length - 1;

    while (index >= 0 && isXmlWhitespaceCode(source.charCodeAt(index))) {
        index -= 1;
    }

    return index >= 0 && source.charCodeAt(index) === 47;
}

export function isXmlWhitespaceCode(code: number): boolean {
    return code === 9 || code === 10 || code === 13 || code === 32;
}

export function parseCellColumnNumberFast(address: string): number {
    let columnNumber = 0;
    let index = 0;

    while (index < address.length) {
        let characterCode = address.charCodeAt(index);
        if (characterCode === 36) {
            index += 1;
            continue;
        }

        if (characterCode >= 97 && characterCode <= 122) {
            characterCode -= 32;
        }

        if (characterCode < 65 || characterCode > 90) {
            break;
        }

        columnNumber = columnNumber * 26 + (characterCode - 64);
        index += 1;
    }

    if (columnNumber === 0) {
        throw new XlsxError(`Invalid cell address: ${address}`);
    }

    return columnNumber;
}

export function readXmlAttrFast(source: string, attributeName: string): string | undefined {
    const pattern = attributeName;
    let searchStart = 0;

    while (searchStart < source.length) {
        const attributeStart = source.indexOf(pattern, searchStart);
        if (attributeStart === -1) {
            return undefined;
        }

        const previousCode = attributeStart === 0 ? 32 : source.charCodeAt(attributeStart - 1);
        if (isXmlAttributeBoundaryCode(previousCode)) {
            let cursor = attributeStart + pattern.length;

            while (cursor < source.length && isXmlWhitespaceCode(source.charCodeAt(cursor))) {
                cursor += 1;
            }

            if (source.charCodeAt(cursor) !== 61) {
                searchStart = attributeStart + pattern.length;
                continue;
            }

            cursor += 1;
            while (cursor < source.length && isXmlWhitespaceCode(source.charCodeAt(cursor))) {
                cursor += 1;
            }

            const quote = source.charCodeAt(cursor);
            if (quote !== 34 && quote !== 39) {
                searchStart = attributeStart + pattern.length;
                continue;
            }

            const valueStart = cursor + 1;
            const valueEnd = source.indexOf(String.fromCharCode(quote), valueStart);
            return valueEnd === -1 ? undefined : source.slice(valueStart, valueEnd);
        }

        searchStart = attributeStart + pattern.length;
    }

    return undefined;
}

export function scanValueWindowCellFast(
    xml: string,
    rowInnerEnd: number,
    cellCursor: number
): ValueWindowCellScanResult | null {
    while (cellCursor < rowInnerEnd) {
        const cellStart = xml.indexOf("<c", cellCursor);
        if (cellStart === -1 || cellStart >= rowInnerEnd) {
            return null;
        }

        const nextCode = xml.charCodeAt(cellStart + 2);
        if (!isTagBoundaryCode(nextCode)) {
            cellCursor = cellStart + 2;
            continue;
        }

        const cellOpenTagEnd = xml.indexOf(">", cellStart + 2);
        if (cellOpenTagEnd === -1 || cellOpenTagEnd > rowInnerEnd) {
            return null;
        }

        const cellMetadata = parseValueWindowCellTagMetadataFast(
            xml,
            cellStart + 2,
            cellOpenTagEnd
        );
        const cellEnd = cellMetadata.selfClosing
            ? cellOpenTagEnd + 1
            : xml.indexOf("</c>", cellOpenTagEnd + 1);
        if (!cellMetadata.addressSource || cellEnd === -1) {
            cellCursor = cellOpenTagEnd + 1;
            continue;
        }

        const innerStart = cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1;
        const innerEnd = cellEnd;
        return {
            addressSource: cellMetadata.addressSource,
            columnNumber: parseCellColumnNumberFast(cellMetadata.addressSource),
            innerEnd,
            innerStart,
            logical: hasLogicalValueWindowCellContent(
                xml,
                innerStart,
                innerEnd,
                cellMetadata.rawType
            ),
            nextCursor: cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length,
            rawType: cellMetadata.rawType,
        };
    }

    return null;
}

function parseValueWindowCellTagMetadataFast(
    xml: string,
    start: number,
    end: number
): { addressSource: string | undefined; rawType: string | null; selfClosing: boolean } {
    let addressSource: string | undefined;
    let rawType: string | null = null;
    let index = start;

    while (index < end) {
        while (index < end && isXmlWhitespaceCode(xml.charCodeAt(index))) {
            index += 1;
        }
        if (index >= end) {
            break;
        }
        if (xml.charCodeAt(index) === 47) {
            break;
        }

        const nameStart = index;
        while (index < end) {
            const code = xml.charCodeAt(index);
            if (code === 61 || code === 47 || isXmlWhitespaceCode(code)) {
                break;
            }
            index += 1;
        }

        const nameLength = index - nameStart;
        while (index < end && isXmlWhitespaceCode(xml.charCodeAt(index))) {
            index += 1;
        }
        if (index >= end || xml.charCodeAt(index) !== 61) {
            while (index < end && !isXmlWhitespaceCode(xml.charCodeAt(index))) {
                index += 1;
            }
            continue;
        }

        index += 1;
        while (index < end && isXmlWhitespaceCode(xml.charCodeAt(index))) {
            index += 1;
        }
        if (index >= end) {
            break;
        }

        const quote = xml.charCodeAt(index);
        if (quote !== 34 && quote !== 39) {
            continue;
        }

        const valueStart = index + 1;
        const valueEnd = xml.indexOf(String.fromCharCode(quote), valueStart);
        if (valueEnd === -1 || valueEnd > end) {
            break;
        }

        if (nameLength === 1) {
            const nameCode = xml.charCodeAt(nameStart);
            if (nameCode === 114) {
                addressSource = xml.slice(valueStart, valueEnd);
            } else if (nameCode === 116) {
                rawType = xml.slice(valueStart, valueEnd);
            }
        }

        index = valueEnd + 1;
    }

    return {
        addressSource,
        rawType,
        selfClosing: isSelfClosingTagRange(xml, start, end),
    };
}

function hasCellChildTagFast(xml: string, start: number, end: number, tagName: string): boolean {
    const pattern = `<${tagName}`;
    let searchStart = start;

    while (searchStart < end) {
        const tagStart = xml.indexOf(pattern, searchStart);
        if (tagStart === -1 || tagStart >= end) {
            return false;
        }

        const boundaryCode = xml.charCodeAt(tagStart + pattern.length);
        if (isTagBoundaryCode(boundaryCode)) {
            return true;
        }

        searchStart = tagStart + pattern.length;
    }

    return false;
}

function isSelfClosingTagRange(xml: string, start: number, end: number): boolean {
    let index = end - 1;

    while (index >= start && isXmlWhitespaceCode(xml.charCodeAt(index))) {
        index -= 1;
    }

    return index >= start && xml.charCodeAt(index) === 47;
}

function isTagBoundaryCode(code: number): boolean {
    return code === 47 || code === 62 || isXmlWhitespaceCode(code);
}

function isXmlAttributeBoundaryCode(code: number): boolean {
    return code === 47 || isXmlWhitespaceCode(code);
}
