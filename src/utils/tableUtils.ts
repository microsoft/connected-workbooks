import JSZip from "jszip";
import { TableData } from "../types";
import {
    InvalidColumnNameErr,
    defaults,
    element,
    elementAttributes,
    queryTableNotFoundErr,
    queryTableXmlPath,
    sheetsNotFoundErr,
    sheetsXmlPath,
    tableNotFoundErr,
    tableXmlPath,
    textResultType,
    workbookXmlPath,
    xmlTextResultType,
} from "./constants";
import documentUtils from "./documentUtils";
import { v4 } from "uuid";

const updateTableInitialDataIfNeeded = async (zip: JSZip, tableData?: TableData, updateQueryTable?: boolean): Promise<void> => {
    if (!tableData) {
        return;
    }

    const sheetsXmlString: string | undefined = await zip.file(sheetsXmlPath)?.async(textResultType);
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const newSheet: string = updateSheetsInitialData(sheetsXmlString, tableData);
    zip.file(sheetsXmlPath, newSheet);

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

const updateSheetsInitialData = (sheetsXmlString: string, tableData: TableData): string => {
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
        columnRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, col));
    });
    sheetData.appendChild(columnRow);
    rowIndex++;
    tableData.rows.forEach((row) => {
        const newRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, element.row);
        newRow.setAttribute(elementAttributes.row, (rowIndex + 1).toString());
        newRow.setAttribute(elementAttributes.spans, "1:" + row.length);
        newRow.setAttribute(elementAttributes.x14acDyDescent, "0.3");
        row.forEach((cellContent, colIndex) => {
            newRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, cellContent));
        });
        sheetData.appendChild(newRow);
        rowIndex++;
    });

    sheetsDoc
        .getElementsByTagName(element.dimension)[0]
        .setAttribute(elementAttributes.reference, documentUtils.getTableReference(tableData.rows[0].length - 1, tableData.rows.length));

    return serializer.serializeToString(sheetsDoc);
};

const getAdjustedColumnNames = (columnNames: (string | number | boolean)[]): string[] => {
    const newColumnNames: string[] = [];
    columnNames.forEach((columnName) => newColumnNames.push(getNextAvailableColumnName(newColumnNames, getColumnNameToString(columnName))));
    return newColumnNames;
};

const getColumnNameToString = (columnName: string | number | boolean): string => {
    if (columnName === null || (typeof columnName === "string" && columnName.length == 0)) {
        return defaults.columnName;
    }

    return columnName.toString();
};

const getNextAvailableColumnName = (columnNames: string[], columnName: string): string => {
    let index = 1;
    let nextAvailableName = columnName;
    while (columnNames.includes(nextAvailableName)) {
        nextAvailableName = `${columnName} (${index})`;
        index++;
    }

    return nextAvailableName;
};

const getRawColumnNames = (columnNames: (string | number | boolean)[]): string[] => {
    const newColumnNames: string[] = [];
    columnNames.forEach((columnName) => newColumnNames.push(getColumnNameOrReiseError(newColumnNames, columnName)));

    return newColumnNames;
};

const getColumnNameOrReiseError = (columnNames: string[], columnName: string | number | boolean): string => {
    // column name shouldn't be empty.
    if (columnName === null || (typeof columnName === "string" && columnName.length == 0)) {
        throw new Error(InvalidColumnNameErr);
    }

    // Duplicate column name.
    if (columnNames.includes(columnName.toString())) {
        throw new Error(InvalidColumnNameErr);
    }

    return columnName.toString();
};

export default {
    updateTableInitialDataIfNeeded,
    updateSheetsInitialData,
    updateWorkbookInitialData,
    updateTablesInitialData,
    updateQueryTablesInitialData,
    getNextAvailableColumnName,
    getAdjustedColumnNames,
    getRawColumnNames,
};
