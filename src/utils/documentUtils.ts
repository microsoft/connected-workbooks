// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import {
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
    // 65 is the ascii value of first column 'A'
    return "$" + String.fromCharCode(col + 65) + "$" + row.toString();
};

const getCellReferenceRelative = (col: number, row: number): string => {
    // 65 is the ascii value of first column 'A'
    return String.fromCharCode(col + 65) + row.toString();
};

const getTableReference = (numberOfCols: number, numberOfRows: number) => {
    return `A1:${getCellReferenceRelative(numberOfCols, numberOfRows)}`;
};

const createCellElement = (doc: Document, colIndex: number, rowIndex: number, data: string) => {
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
    cellData.textContent = data;
};

const resolveType = (originalData: string | number | boolean, rowHeader: boolean): DataTypes => {
    const data: string = originalData as string;
    if (rowHeader) {
        // Headers should be string by default.
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
    resolveType,
};
