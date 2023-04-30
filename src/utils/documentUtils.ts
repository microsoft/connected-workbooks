// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { docPropsCoreXmlPath, docPropsRootElement } from "../constants";
import { dataTypes } from "../types";

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
        elements[0]!.innerHTML = value!;
    }
};

const getDocPropsProperties = async (zip: JSZip): Promise<{ doc: Document; properties: Element }> => {
    const docPropsCoreXmlString = await zip.file(docPropsCoreXmlPath)?.async("text");
    if (docPropsCoreXmlString === undefined) {
        throw new Error("DocProps core.xml was not found in template");
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(docPropsCoreXmlString, "text/xml");

    const properties = doc.getElementsByTagName(docPropsRootElement).item(0);
    if (properties === null) {
        throw new Error("Invalid DocProps core.xml");
    }

    return { doc, properties };
};

const getCellReference = (col: number, row: number): string => {
    // 65 is the ascii value of first column 'A'
    return String.fromCharCode(col + 65) + "$" + row.toString();
};

const getTableReference = (numberOfCols: number, numberOfRows: number) => {
    return `A1:${getCellReference(numberOfCols, numberOfRows)}`.replace("$", "");
};

const createCellElement = (doc: Document, colIndex: number, rowIndex: number, dataType: number, data: string) => {
    const cell = doc.createElementNS(doc.documentElement.namespaceURI, "c");
    cell.setAttribute("r", getCellReference(colIndex, rowIndex + 1).replace("$", ""));
    const cellData = doc.createElementNS(doc.documentElement.namespaceURI, "v");
    updateCellData(dataType, data, cell, cellData);
    cell.appendChild(cellData);
    return cell;
};

const updateCellData = (dataType: number, data: string, cell: Element, cellData: Element) => {
    switch(dataType) {
    case dataTypes.string:
        cell.setAttribute("t", "str");
        break;
    case dataTypes.number:
        cell.setAttribute("t", "1");
        break;
    case dataTypes.boolean:
        cell.setAttribute("t", "b");
        break;
    }
    cellData.textContent = data;
};

export default { createOrUpdateProperty, getDocPropsProperties, getCellReference, createCell: createCellElement, getTableReference };
