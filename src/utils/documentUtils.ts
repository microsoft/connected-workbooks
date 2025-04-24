// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import {
    columnIndexOutOfRangeErr,
    dataTypeKind,
    docPropsCoreXmlPath,
    docPropsRootElement,
    element,
    elementAttributes,
    falseStr,
    textResultType,
    trueStr,
    xmlTextResultType,
} from "./constants";
import { DataTypes } from "../types";
import { DOMParser } from "xmldom-qsa";

const createOrUpdateProperty = (doc: Document, parent: Element, property: string, value?: string | null): void => {
    if (value === undefined) {
        return;
    }

    const elements = parent.getElementsByTagName(property);

    if (elements?.length === 0) {
        const newElement = doc.createElement(property);
        newElement.textContent = value;
        parent.appendChild(newElement);
    } else if (elements.length > 1) {
        throw new Error(`Invalid DocProps core.xml, multiple ${property} elements`);
    } else if (elements?.length > 0) {
        elements[0]!.textContent = value;
    }
};

const getDocPropsProperties = async (zip: JSZip): Promise<{ doc: Document; properties: Element }> => {
    const docPropsCoreXmlString = await zip.file(docPropsCoreXmlPath)?.async(textResultType);
    if (docPropsCoreXmlString === undefined) {
        throw new Error("DocProps core.xml was not found in template");
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(docPropsCoreXmlString, xmlTextResultType);

    const properties = doc.getElementsByTagName(docPropsRootElement).item(0);
    if (properties === null) {
        throw new Error("Invalid DocProps core.xml");
    }

    return { doc, properties };
};

const getCellReferenceAbsolute = (col: number, row: number): string => {
    return "$" + convertToExcelColumn(col) + "$" + row.toString();
};

const getCellReferenceRelative = (col: number, row: number): string => {
    return convertToExcelColumn(col) + row.toString();
};

const convertToExcelColumn = (index: number): string => {
    if (index >= 16384) {
        throw new Error(columnIndexOutOfRangeErr);
    }

    let columnStr = "";
    const base = 26; // number of letters in the alphabet
    while (index >= 0) {
        const remainder = index % base;
        columnStr = String.fromCharCode(remainder + 65) + columnStr; // ASCII 'A' is 65
        index = Math.floor(index / base) - 1;
    }

    return columnStr;
};

const getTableReference = (numberOfCols: number, numberOfRows: number,startCol: number, startRow: number): string => {
    return `${getCellReferenceRelative(startCol, startRow)}:${getCellReferenceRelative(numberOfCols, numberOfRows)}`;
};

const createCellElement = (doc: Document, colIndex: number, rowIndex: number, data: string): Element => {
    const cell: Element = doc.createElementNS(doc.documentElement.namespaceURI, element.kindCell);
    cell.setAttribute(elementAttributes.row, getCellReferenceRelative(colIndex, rowIndex + 1));
    const cellData: Element = doc.createElementNS(doc.documentElement.namespaceURI, element.cellValue);
    updateCellData(data, cell, cellData, rowIndex === 0);
    cell.appendChild(cellData);

    return cell;
};

const updateCellData = (data: string, cell: Element, cellData: Element, rowHeader: boolean) => {
    switch (resolveType(data, rowHeader)) {
        case DataTypes.string:
            cell.setAttribute(element.text, dataTypeKind.string);
            break;
        case DataTypes.number:
            cell.setAttribute(element.text, dataTypeKind.number);
            break;
        case DataTypes.boolean:
            cell.setAttribute(element.text, dataTypeKind.boolean);
            break;
    }
    if (data.startsWith(" ") || data.endsWith(" ")) {
        cellData.setAttribute(elementAttributes.space, "preserve");
    }

    cellData.textContent = data;
};

const resolveType = (originalData: string | number | boolean, rowHeader: boolean): DataTypes => {
    const data: string = originalData as string;
    if (rowHeader || data.trim() === "") {
        // Headers and whitespace should be string by default.
        return DataTypes.string;
    }
    let dataType: DataTypes = isNaN(Number(data)) ? DataTypes.string : DataTypes.number;
    if (dataType == DataTypes.string) {
        if (data.trim() == trueStr || data.trim() == falseStr) {
            dataType = DataTypes.boolean;
        }
    }

    return dataType;
};

export default {
    createOrUpdateProperty,
    getDocPropsProperties,
    getCellReferenceRelative,
    getCellReferenceAbsolute,
    createCell: createCellElement,
    getTableReference,
    updateCellData,
    resolveType,
    convertToExcelColumn,
};
