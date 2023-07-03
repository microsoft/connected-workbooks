// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { defaults, gridNotFoundErr } from "../utils/constants";
import { Grid, TableData } from "../types";
import columnNameUtils from "./columnNameUtils";

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
        return columnNameUtils.getAdjustedColumnNames(grid.data[0]);
    }

    // Get column names and fails if it's not a legal name.
    return columnNameUtils.getRawColumnNames(grid.data[0]);
};

export default { parseToTableData };
