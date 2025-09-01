// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { TableData } from "../types";
import {
    element,
    elementAttributes,
    invalidCellValueErr,
    maxCellCharacters,
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
 * @param cellRangeRef - Cell range reference (e.g. "A1:C5").
 * @param sheetPath - Path to the sheet XML within the zip.
 * @param tablePath - Path to the table XML within the zip.
 * @param tableName - Name of the table.
 * @param tableData - Optional TableData containing headers and rows.
 * @param updateQueryTable - Whether to update the associated queryTable part.
 */
const updateTableInitialDataIfNeeded = async (zip: JSZip, cellRangeRef: string, sheetPath: string, tablePath: string, sheetName: string, tableData?: TableData, updateQueryTable?: boolean): Promise<void> => {
    if (!tableData) {
        return;
    }

    const sheetsXmlString: string | undefined = await zip.file(sheetPath)?.async(textResultType);
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const newSheet: string = updateSheetsInitialData(sheetsXmlString, tableData, cellRangeRef);
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

        const newWorkbook: string = updateWorkbookInitialData(workbookXmlString, sheetName + GenerateReferenceFromString(cellRangeRef));
        zip.file(workbookXmlPath, newWorkbook);
    }

    const tableXmlString: string | undefined = await zip.file(tablePath)?.async(textResultType);
    if (tableXmlString === undefined) {
        throw new Error(tableNotFoundErr);
    }

    const newTable: string = updateTablesInitialData(tableXmlString, tableData, cellRangeRef, updateQueryTable);
    zip.file(tablePath, newTable);
};

/**
 * Generate updated table XML string with new columns, reference, and filter range.
 * @param tableXmlString - Original table XML.
 * @param tableData - TableData containing column names.
 * @param cellRangeRef - Cell range reference.
 * @param updateQueryTable - Whether to include queryTable attributes.
 * @returns Serialized XML string of the updated table.
 */
const updateTablesInitialData = (tableXmlString: string, tableData: TableData, cellRangeRef: string, updateQueryTable = false): string => {
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
        .setAttribute(elementAttributes.reference, cellRangeRef);
    tableDoc
        .getElementsByTagName(element.autoFilter)[0]
        .setAttribute(elementAttributes.reference, cellRangeRef);

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
    definedName.textContent = customDefinedName;

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
 * @param cellRangeRef - Cell range reference.
 * @returns Serialized XML string of the updated sheet.
 */
const updateSheetsInitialData = (sheetsXmlString: string, tableData: TableData, cellRangeRef: string): string => {
    let { row, column } = documentUtils.GetStartPosition(cellRangeRef);
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, xmlTextResultType);
    const sheetData: Element = sheetsDoc.getElementsByTagName(element.sheetData)[0];
    sheetData.textContent = "";

    const columnRow: Element = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
    columnRow.setAttribute(elementAttributes.row, row.toString());
    columnRow.setAttribute(elementAttributes.spans, column + ":" + (column + tableData.columnNames.length - 1));
    columnRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
    tableData.columnNames.forEach((col: string, colIndex: number) => {
<<<<<<< HEAD
        validateCellContentLength(col);
        columnRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, col));
=======
        columnRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex + column - 1, row - 1, col));
>>>>>>> main
    });
    sheetData.appendChild(columnRow);
    row++;

    tableData.rows.forEach((_row) => {
        const newRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
        newRow.setAttribute(elementAttributes.row, row.toString());
        newRow.setAttribute(elementAttributes.spans, column + ":" + (column + tableData.columnNames.length - 1));
        newRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
<<<<<<< HEAD
        row.forEach((cellContent, colIndex) => {
            validateCellContentLength(cellContent);
            newRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, cellContent));
=======
        _row.forEach((cellContent, colIndex) => {
            newRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex + column - 1, row - 1, cellContent));
>>>>>>> main
        });
        sheetData.appendChild(newRow);
        row++;
    });

    sheetsDoc.getElementsByTagName(element.dimension)[0].setAttribute(elementAttributes.reference, cellRangeRef);
    sheetsDoc.getElementsByTagName(element.selection)[0].setAttribute(elementAttributes.sqref, cellRangeRef);
    return serializer.serializeToString(sheetsDoc);
};

<<<<<<< HEAD
const validateCellContentLength = (cellContent: string): void => {
    if (cellContent.length > maxCellCharacters) {
        throw new Error(invalidCellValueErr);
    }
=======
/**
 * Add Excel-style dollar signs and a '!' prefix to a cell range.
 * Converts "A1:B2" into "!$A$1:$B$2".
 * @param cellRangeRef - Range reference string without dollar signs.
 * @returns Range with dollar signs and prefix.
 */
const GenerateReferenceFromString = (cellRangeRef: string): string => {
    return "!" + cellRangeRef.split(":").map(part => {
        const match = part.match(/^([A-Za-z]+)(\d+)$/);
        if (match) {
            const [, col, row] = match;
            return `$${col.toUpperCase()}$${row}`;
        }
    }).join(":");
>>>>>>> main
}

export default {
    updateTableInitialDataIfNeeded,
    updateSheetsInitialData,
    updateWorkbookInitialData,
    updateTablesInitialData,
    updateQueryTablesInitialData,
    GenerateReferenceFromString,
};
