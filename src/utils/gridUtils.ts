// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { defaults, gridNotFoundErr, InvalidColumnNameErr } from "../utils/constants";
import { Grid, TableData } from "../types";

const parseToTableData = (grid: Grid): TableData => {
    const columnNames: string[] = generateColumnNames(grid);
    const rows: string[][] = parseGridRows(grid);

    return { columnNames: columnNames, rows: rows };
};

const parseGridRows = (grid: Grid): string[][] => {
    const gridData: (string | number | boolean)[][] = grid.data;
    if (!gridData) {
        throw new Error(gridNotFoundErr);
    }

    const rows: string[][] = [];
    if (!grid.config?.promoteHeaders) {
        const row: string[] = [];
        for (const prop in gridData[0]) {
            const cellValue: string | number | boolean = gridData[0][prop];
            row.push(cellValue.toString());
        }

        rows.push(row);
    }

    for (let i = 1; i < gridData.length; i++) {
        const rowData: (string | number | boolean)[] = gridData[i];
        const row: string[] = [];
        for (const prop in rowData) {
            const cellValue: string | number | boolean = rowData[prop];
            row.push(cellValue?.toString() ?? "");
        }

        rows.push(row);
    }

    return rows;
};

const generateColumnNames = (grid: Grid): string[] => {
    const columnNames: string[] = [];
    if (!grid.config?.promoteHeaders) {
        for (let i = 0; i < grid.data[0].length; i++) {
            columnNames.push(`${defaults.columnName} ${i + 1}`);
        }

        return columnNames;
    }

    // We are adjusting column names by default.
    if (!grid.config || grid.config.adjustColumnNames === undefined || grid.config.adjustColumnNames) {
        return getAdjustedColumnNames(grid.data[0]);
    }

    // Get column names and fails if it's not a legal name.
    return getRawColumnNames(grid.data[0]);
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
    columnNames.forEach((columnName) => newColumnNames.push(getColumnNameOrRaiseError(newColumnNames, columnName)));

    return newColumnNames;
};

const getColumnNameOrRaiseError = (columnNames: string[], columnName: string | number | boolean): string => {
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

export default { parseToTableData };
