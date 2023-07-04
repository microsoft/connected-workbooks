// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { defaults, InvalidColumnNameErr } from "../utils/constants";
import { Grid, TableData } from "../types";

interface MergedGridConfig {
    promoteHeaders: boolean;
    adjustColumnNames: boolean;
}
interface MergedGrid {
    data: string[][];
    config: MergedGridConfig;
}

const parseToTableData = (grid: Grid): TableData => {
    const mergedGrid: MergedGrid = {
        config: {
            promoteHeaders: grid.config?.promoteHeaders ?? false,
            adjustColumnNames: grid.config?.adjustColumnNames ?? true,
        },
        data: grid.data.map((row) => row.map((value) => value.toString())),
    };

    validateGrid(mergedGrid);
    const columnNames: string[] = generateColumnNames(mergedGrid);

    if (mergedGrid.config.promoteHeaders) {
        mergedGrid.data.shift();
    }
    return { columnNames: columnNames, rows: mergedGrid.data };
};

/*
 * Validates the grid, throws an error if the grid is invalid.
 * A valid grid has:
 * - MxN structure.
 * - If promoteHeaders is true - has at least 1 row, and in case adjustColumnNames is false, first row is unique and non empty.
 */
const validateGrid = (grid: MergedGrid): void => {
    if (!validateDataArrayDimensions(grid.data)) {
        throw new Error("Invalid grid dimensions");
    }

    if (grid.config.promoteHeaders && grid.data.length === 0) {
        throw new Error("Promote headers is not supported for an empty grid");
    }

    if (grid.config.promoteHeaders && grid.config.adjustColumnNames === false && !validateUniqueAndValidDataArray(grid.data[0])) {
        throw new Error("Headers cannot be promoted without adjusting column names");
    }
};

const validateDataArrayDimensions = (arr: unknown[][]): boolean => {
    if (arr.length === 0) {
        return true; // Empty array is considered valid
    }
    const innerLength = arr[0].length;

    if (innerLength === 0) {
        return false; // [[]] and any [] innerArr is invalid
    }

    return arr.every((innerArr) => innerArr.length === innerLength);
};

const validateUniqueAndValidDataArray = (arr: string[]): boolean => {
    if (arr.some((element) => element === "")) {
        return false; // Array contains empty elements
    }

    const uniqueSet = new Set(arr);
    return uniqueSet.size === arr.length;
};

const generateColumnNames = (grid: MergedGrid): string[] => {
    const columnNames: string[] = [];
    if (!grid.config.promoteHeaders) {
        for (let i = 0; i < grid.data[0].length; i++) {
            columnNames.push(`${defaults.columnName} ${i + 1}`);
        }

        return columnNames;
    }

    if (grid.config.adjustColumnNames) {
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
