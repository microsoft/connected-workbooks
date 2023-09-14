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
    invalidDateTimeErr,
    textResultType,
    trueStr,
    xmlTextResultType,
} from "./constants";
import { DataTypes } from "../types";
import dateTimeUtils from "./dateTimeUtils";

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

const getTableReference = (numberOfCols: number, numberOfRows: number): string => {
    return `A1:${getCellReferenceRelative(numberOfCols, numberOfRows)}`;
};

const getFormatStyleIndex = (dateTimeFormat: DataTypes, formats: DataTypes[]) => {
    if (dateTimeFormat < DataTypes.shortTime || dateTimeFormat > DataTypes.longDate) {
        throw new Error(invalidDateTimeErr);
    }

    // Style Index is the index of the format in the styleFormats array + 1 
    const styleIndex: number = formats.includes(dateTimeFormat) ? formats.indexOf(dateTimeFormat) + 1 : formats.push(dateTimeFormat);
    return styleIndex.toString();
 };

const createCellElement = (doc: Document, colIndex: number, rowIndex: number, data: string, formats: DataTypes[]): Element => {
    const cell: Element = doc.createElementNS(doc.documentElement.namespaceURI, element.kindCell);
    cell.setAttribute(elementAttributes.row, getCellReferenceRelative(colIndex, rowIndex + 1));
    const cellData: Element = doc.createElementNS(doc.documentElement.namespaceURI, element.cellValue);
    updateCellData(data, cell, cellData, rowIndex === 0, formats);
    cell.appendChild(cellData);

    return cell;
};

const updateCellData = (data: string, cell: Element, cellData: Element, rowHeader: boolean, formats: DataTypes[]) => {
    let dataType: DataTypes = resolveType(data, rowHeader);
    switch (dataType) {
        case DataTypes.string:
            cell.setAttribute(element.text, dataTypeKind.string);
            break;
        case DataTypes.number:
            cell.setAttribute(element.text, dataTypeKind.number);
            break;
        case DataTypes.boolean:
            cell.setAttribute(element.text, dataTypeKind.boolean);
            break;
        // All other data types are datetimes
        default:
            cell.setAttribute(elementAttributes.style, getFormatStyleIndex(dataType, formats));
            data = dateTimeUtils.convertToExcelDate(data, dataType).toString();
    }
    
    if (dataType == DataTypes.string && (data.startsWith(" ") || data.endsWith(" "))) {
        cellData.setAttribute(elementAttributes.space, "preserve");        
    }

    cellData.textContent = data;
};

const resolveType = (originalData: string | number | boolean, rowHeader: boolean): DataTypes => {
    const data: string = originalData as string;
    const dateTimeFormat: DataTypes|undefined = dateTimeUtils.detectDateTimeFormat(data);
    if (dateTimeFormat != undefined) {
            return dateTimeFormat;
    }

    if ((rowHeader) || (data.trim() === "")) {
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
};
