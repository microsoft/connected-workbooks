// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { DataTypes, TableData } from "../types";
import {
    customNumberFormatId,
    defaults,
    element,
    elementAttributes,
    falseValue,
    queryTableNotFoundErr,
    queryTableXmlPath,
    sheetsNotFoundErr,
    sheetsXmlPath,
    shortDateId,
    stylesNotFoundErr,
    stylesXmlPath,
    tableNotFoundErr,
    tableXmlPath,
    textResultType,
    trueValue,
    workbookXmlPath,
    xmlTextResultType,
} from "./constants";
import documentUtils from "./documentUtils";
import { v4 } from "uuid";
import dateTimeUtils, { DateTimeFormat } from "./dateTimeUtils";

const updateTableInitialDataIfNeeded = async (zip: JSZip, tableData?: TableData, updateQueryTable?: boolean): Promise<void> => {
    if (!tableData) {
        return;
    }

    const sheetsXmlString: string | undefined = await zip.file(sheetsXmlPath)?.async(textResultType);
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const {newSheet, formatStyles} = updateSheetsInitialData(sheetsXmlString, tableData);
    zip.file(sheetsXmlPath, newSheet);

    if (formatStyles.length > 0) {
        const stylesXmlString: string | undefined = await zip.file(stylesXmlPath)?.async(textResultType);
        if (stylesXmlString === undefined) {
            throw new Error(stylesNotFoundErr);
        }

        const newStyles: string = await updateStylesInitialData(stylesXmlString, formatStyles);
        zip.file(stylesXmlPath, newStyles);
    }

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

        const newWorkbook: string = updateWorkbookInitialData(workbookXmlString, tableData);
        zip.file(workbookXmlPath, newWorkbook);
    }

    const tableXmlString: string | undefined = await zip.file(tableXmlPath)?.async(textResultType);
    if (tableXmlString === undefined) {
        throw new Error(tableNotFoundErr);
    }

    const newTable: string = updateTablesInitialData(tableXmlString, tableData, updateQueryTable);
    zip.file(tableXmlPath, newTable);
};

const updateStylesInitialData = (stylesXmlString: string, formatStyles: DateTimeFormat[]): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const stylesDoc: Document = parser.parseFromString(stylesXmlString, xmlTextResultType);
    let numFmtsCount: number = 0;
    const cellXfs: Element = stylesDoc.getElementsByTagName(element.cellXfs)[0];
    const fonts: Element = stylesDoc.getElementsByTagName(element.fonts)[0];
    const numFmts: Element = stylesDoc.createElementNS(stylesDoc.documentElement.namespaceURI, element.numFmts);
    formatStyles.forEach((dateTimeFormat, i) => {
        // Add new numfmtId when necessary
        let numFmtId: number = dateTimeFormat.formatId ? dateTimeFormat.formatId : i + customNumberFormatId;
        if (dateTimeFormat.formatCode) {
            const numFmt: Element = stylesDoc.createElementNS(stylesDoc.documentElement.namespaceURI, element.numFmt);
            numFmt.setAttribute(elementAttributes.numFmtId, numFmtId.toString());
            numFmt.setAttribute(elementAttributes.formatCode, dateTimeFormat.formatCode);
            numFmts.appendChild(numFmt);
            numFmtsCount++;
        }
        
        const xf: Element = stylesDoc.createElementNS(stylesDoc.documentElement.namespaceURI, element.xf);
        xf.setAttribute(elementAttributes.numFmtId, numFmtId.toString());
        xf.setAttribute(elementAttributes.fontId, falseValue);
        xf.setAttribute(elementAttributes.fillId, falseValue);
        xf.setAttribute(elementAttributes.borderId, falseValue);
        xf.setAttribute(elementAttributes.xfId, falseValue);
        xf.setAttribute(elementAttributes.applyNumberFormat, trueValue);
        cellXfs.appendChild(xf);
    });
    // count includes the formatStyles and default null format
    if (numFmtsCount > 0) {
        stylesDoc.getElementsByTagName(element.styleSheet)[0].insertBefore(numFmts, fonts);
        numFmts.setAttribute(elementAttributes.count, numFmtsCount.toString());
    }

    return serializer.serializeToString(stylesDoc);
};

const updateTablesInitialData = (tableXmlString: string, tableData: TableData, updateQueryTable = false): string => {
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
        .setAttribute(elementAttributes.reference, `A1:${documentUtils.getCellReferenceRelative(tableData.columnNames.length - 1, tableData.rows.length + 1)}`);
    tableDoc
        .getElementsByTagName(element.autoFilter)[0]
        .setAttribute(elementAttributes.reference, `A1:${documentUtils.getCellReferenceRelative(tableData.columnNames.length - 1, tableData.rows.length + 1)}`);

    return serializer.serializeToString(tableDoc);
};

const updateWorkbookInitialData = (workbookXmlString: string, tableData: TableData): string => {
    const newParser: DOMParser = new DOMParser();
    const newSerializer: XMLSerializer = new XMLSerializer();
    const workbookDoc: Document = newParser.parseFromString(workbookXmlString, xmlTextResultType);
    const definedName: Element = workbookDoc.getElementsByTagName(element.definedName)[0];
    definedName.textContent =
        defaults.sheetName + `!$A$1:${documentUtils.getCellReferenceAbsolute(tableData.columnNames.length - 1, tableData.rows.length + 1)}`;

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

const updateSheetsInitialData = (sheetsXmlString: string, tableData: TableData) => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, xmlTextResultType);
    const sheetData: Element = sheetsDoc.getElementsByTagName(element.sheetData)[0];
    sheetData.textContent = "";
    let rowIndex = 0;
    const columnRow: Element = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
    columnRow.setAttribute(elementAttributes.row, (rowIndex + 1).toString());
    columnRow.setAttribute(elementAttributes.spans, "1:" + tableData.columnNames.length);
    columnRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
    tableData.columnNames.forEach((col: string, colIndex: number) => {
        columnRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, col, []));
    });
    sheetData.appendChild(columnRow);
    let formatStyles: DateTimeFormat[] = [];
    rowIndex++;
    tableData.rows.forEach((row) => {
        const newRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
        newRow.setAttribute(elementAttributes.row, (rowIndex + 1).toString());
        newRow.setAttribute(elementAttributes.spans, "1:" + row.length);
        newRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
        row.forEach((cellContent, colIndex) => {
            newRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, cellContent, formatStyles));
        });
        sheetData.appendChild(newRow);
        rowIndex++;
    });
    const reference = documentUtils.getTableReference(tableData.rows[0].length - 1, tableData.rows.length + 1);

    sheetsDoc.getElementsByTagName(element.dimension)[0].setAttribute(elementAttributes.reference, reference);
    sheetsDoc.getElementsByTagName(element.selection)[0].setAttribute(elementAttributes.sqref, reference);
    const newSheet = serializer.serializeToString(sheetsDoc);
    return {newSheet, formatStyles};
};

export default {
    updateTableInitialDataIfNeeded,
    updateSheetsInitialData,
    updateWorkbookInitialData,
    updateTablesInitialData,
    updateQueryTablesInitialData,
};
