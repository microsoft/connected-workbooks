// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { TableData } from "../types";
import {
    element,
    elementAttributes,
    queryTableNotFoundErr,
    queryTableXmlPath,
    sheetsNotFoundErr,
    tableNotFoundErr,
    textResultType,
    workbookXmlPath,
    xmlTextResultType,
} from "./constants";
import documentUtils from "./documentUtils";
import { v4 } from "uuid";
import { DOMParser, XMLSerializer } from "xmldom-qsa";

/**
 * Update initial data for a table, its sheet, query table, and defined name if provided.
 * @param zip - The JSZip instance containing workbook parts.
 * @param ref - Cell range reference (e.g. "A1:C5").
 * @param sheetPath - Path to the sheet XML within the zip.
 * @param tablePath - Path to the table XML within the zip.
 * @param tableName - Name of the table.
 * @param tableData - Optional TableData containing headers and rows.
 * @param updateQueryTable - Whether to update the associated queryTable part.
 */
const updateTableInitialDataIfNeeded = async (zip: JSZip, ref: string, sheetPath: string, tablePath: string, tableName: string, tableData?: TableData, updateQueryTable?: boolean): Promise<void> => {
    if (!tableData) {
        return;
    }

    const sheetsXmlString: string | undefined = await zip.file(sheetPath)?.async(textResultType); // here the correct path sheet
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const newSheet: string = updateSheetsInitialData(sheetsXmlString, tableData, ref);
    zip.file(sheetPath, newSheet);

    if (updateQueryTable) {
        const queryTableXmlString: string | undefined = await zip.file(queryTableXmlPath)?.async(textResultType);
        if (queryTableXmlString === undefined) {
            throw new Error(queryTableNotFoundErr);
        }

        const newQueryTable: string = await updateQueryTablesInitialData(queryTableXmlString, tableData);
        zip.file(queryTableXmlPath, newQueryTable);

        // update defined name
        const workbookXmlString: string | undefined = await zip.file(workbookXmlPath)?.async(textResultType);
        if (workbookXmlString === undefined) {
            throw new Error(sheetsNotFoundErr);
        }

        const newWorkbook: string = updateWorkbookInitialData(workbookXmlString, tableName + addDollar(ref));
        zip.file(workbookXmlPath, newWorkbook);
    }

    const tableXmlString: string | undefined = await zip.file(tablePath)?.async(textResultType);
    if (tableXmlString === undefined) {
        throw new Error(tableNotFoundErr);
    }

    const newTable: string = updateTablesInitialData(tableXmlString, tableData, ref, updateQueryTable);
    zip.file(tablePath, newTable);
};

/**
 * Generate updated table XML string with new columns, reference, and filter range.
 * @param tableXmlString - Original table XML.
 * @param tableData - TableData containing column names.
 * @param ref - Cell range reference.
 * @param updateQueryTable - Whether to include queryTable attributes.
 * @returns Serialized XML string of the updated table.
 */
const updateTablesInitialData = (tableXmlString: string, tableData: TableData, ref: string, updateQueryTable = false): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const tableDoc: Document = parser.parseFromString(tableXmlString, xmlTextResultType);
    const tableColumns: Element = tableDoc.getElementsByTagName(element.tableColumns)[0];
    tableColumns.textContent = "";
    tableData.columnNames.forEach((column: string, columnIndex: number) => {
        const tableColumn: Element = tableDoc.createElementNS(tableDoc.documentElement.namespaceURI, element.tableColumn);
        tableColumn.setAttribute(elementAttributes.id, (columnIndex + 1).toString());
        tableColumn.setAttribute(elementAttributes.name, column);
        tableColumns.appendChild(tableColumn);
        tableColumn.setAttribute(elementAttributes.xr3uid, "{" + v4().toUpperCase() + "}");

        if (updateQueryTable) {
            tableColumn.setAttribute(elementAttributes.uniqueName, (columnIndex + 1).toString());
            tableColumn.setAttribute(elementAttributes.queryTableFieldId, (columnIndex + 1).toString());
        }
    });

    tableColumns.setAttribute(elementAttributes.count, tableData.columnNames.length.toString());
    tableDoc
        .getElementsByTagName(element.table)[0]
        .setAttribute(elementAttributes.reference, ref);
    tableDoc
        .getElementsByTagName(element.autoFilter)[0]
        .setAttribute(elementAttributes.reference, ref);

    return serializer.serializeToString(tableDoc);
};

/**
 * Update the definedName element in workbook XML to a custom name.
 * @param workbookXmlString - Original workbook XML string.
 * @param customDefinedName - New defined name text content (e.g. "!$A$1:$C$5").
 * @returns Serialized XML string of the updated workbook.
 */
const updateWorkbookInitialData = (workbookXmlString: string, customDefinedName: string): string => {
    const newParser: DOMParser = new DOMParser();
    const newSerializer: XMLSerializer = new XMLSerializer();
    const workbookDoc: Document = newParser.parseFromString(workbookXmlString, xmlTextResultType);
    const definedName: Element = workbookDoc.getElementsByTagName(element.definedName)[0];
    definedName.textContent = customDefinedName

    return newSerializer.serializeToString(workbookDoc);
};

const updateQueryTablesInitialData = (queryTableXmlString: string, tableData: TableData): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const queryTableDoc: Document = parser.parseFromString(queryTableXmlString, xmlTextResultType);
    const queryTableFields: Element = queryTableDoc.getElementsByTagName(element.queryTableFields)[0];
    queryTableFields.textContent = "";
    tableData.columnNames.forEach((column: string, columnIndex: number) => {
        const queryTableField: Element = queryTableDoc.createElementNS(queryTableDoc.documentElement.namespaceURI, element.queryTableField);
        queryTableField.setAttribute(elementAttributes.id, (columnIndex + 1).toString());
        queryTableField.setAttribute(elementAttributes.name, column);
        queryTableField.setAttribute(elementAttributes.tableColumnId, (columnIndex + 1).toString());
        queryTableFields.appendChild(queryTableField);
    });
    queryTableFields.setAttribute(elementAttributes.count, tableData.columnNames.length.toString());
    queryTableDoc.getElementsByTagName(element.queryTableRefresh)[0].setAttribute(elementAttributes.nextId, (tableData.columnNames.length + 1).toString());

    return serializer.serializeToString(queryTableDoc);
};

/**
 * Update sheet XML with header row and data rows based on TableData.
 * @param sheetsXmlString - Original sheet XML string.
 * @param tableData - TableData containing headers and rows.
 * @param ref - Cell range reference.
 * @returns Serialized XML string of the updated sheet.
 */
const updateSheetsInitialData = (sheetsXmlString: string, tableData: TableData, ref: string): string => {
    const { row, column } = getRowAndColFromRange(ref);
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, xmlTextResultType);
    const sheetData: Element = sheetsDoc.getElementsByTagName(element.sheetData)[0];
    sheetData.textContent = "";
    let rowIndex = row;

    const columnRow: Element = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
    columnRow.setAttribute(elementAttributes.row, rowIndex.toString());
    columnRow.setAttribute(elementAttributes.spans, column + ":" + (column + tableData.columnNames.length - 1));
    columnRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
    tableData.columnNames.forEach((col: string, colIndex: number) => {
        columnRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex + column - 1, rowIndex - 1, col));
    });
    sheetData.appendChild(columnRow);
    rowIndex++;

    tableData.rows.forEach((_row) => {
        const newRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
        newRow.setAttribute(elementAttributes.row, rowIndex.toString());
        columnRow.setAttribute(elementAttributes.spans, column + ":" + (column + tableData.columnNames.length - 1));
        newRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
        _row.forEach((cellContent, colIndex) => {
            newRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex + column - 1, rowIndex - 1, cellContent));
        });
        sheetData.appendChild(newRow);
        rowIndex++;
    });

    sheetsDoc.getElementsByTagName(element.dimension)[0].setAttribute(elementAttributes.reference, ref);
    sheetsDoc.getElementsByTagName(element.selection)[0].setAttribute(elementAttributes.sqref, ref);
    return serializer.serializeToString(sheetsDoc);
};

/**
 * Parse an Excel range (e.g. "B2:D10") and return its starting row and column indices.
 * @param ref - Range reference string.
 * @returns Object with numeric row and column.
 */
const getRowAndColFromRange = (ref: string): { row: number; column: number } =>{
    const match = ref.match(/^([A-Z]+)(\d+):/);
    if (!match) {
        throw new Error("Invalid range reference format");
    }

    const [, colLetters, rowStr] = match;
    const row = parseInt(rowStr, 10);
    const column = colLetters
        .split("")
        .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - "A".charCodeAt(0) + 1), 0);

    return { row, column };
}

/**
 * Add Excel-style dollar signs and a '!' prefix to a cell range.
 * Converts "A1:B2" into "!$A$1:$B$2".
 * @param ref - Range reference string without dollar signs.
 * @returns Range with dollar signs and prefix.
 */
const addDollar = (ref: string): string => {
    return "!" + ref.split(":").map(part => {
        const match = part.match(/^([A-Za-z]+)(\d+)$/);
        if (match) {
            const [, col, row] = match;
            return `$${col.toUpperCase()}$${row}`;
        }
    }).join(":");
}

export default {
    updateTableInitialDataIfNeeded,
    updateSheetsInitialData,
    updateWorkbookInitialData,
    updateTablesInitialData,
    updateQueryTablesInitialData,
};
