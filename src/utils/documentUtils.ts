// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { docPropsCoreXmlPath, docPropsRootElement } from "../constants";
import { dataTypes } from "../types";

const A:number = 65;

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
    return String.fromCharCode(col + A) + "$" + (row).toString();
} 

const createCell = (doc: Document, parent: Element, colIndex: number, rowIndex: number, dataType: number, data: string): void => {
    const cell = doc.createElementNS(doc.documentElement.namespaceURI, "c");
    cell.setAttribute("r", getCellReference(colIndex, rowIndex + 1).replace("$", ''));
    const cellData = doc.createElementNS(doc.documentElement.namespaceURI, "v");
    updateCellData(dataType, data , cell, cellData);
    cell.appendChild(cellData);
    parent.appendChild(cell);
    colIndex++;
}

const updateCellData = (dataType: number, data: string, newCell: Element, cellData: Element) => {
    if (dataType == dataTypes.string) {
        newCell.setAttribute("t", "str");
        cellData.textContent = data;
    }
    else {
        if (dataType == dataTypes.number) {          
            if (isNaN(Number(data))) {
                data = "0";
            }
            newCell.setAttribute("t", "1");
            cellData.textContent = data;
        }

        if (dataType == dataTypes.boolean) {
            if ((data != "1") && (data != "0")) {
                data = "0";
            }

            newCell.setAttribute("t", "b");
            cellData.textContent = data;
        }
    }
}


export default { createOrUpdateProperty, getDocPropsProperties, getCellReference, createCell };
