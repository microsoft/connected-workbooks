// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { arrayIsntMxNErr, defaults, promotedHeadersCannotBeUsedWithoutAdjustingColumnNamesErr, unexpectedErr } from "../utils/constants";
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

    correctGrid(mergedGrid);
    validateGrid(mergedGrid);

    let columnNames: string[] = [];
    if (mergedGrid.config.promoteHeaders && mergedGrid.config.adjustColumnNames) {
        columnNames = getAdjustedColumnNames(mergedGrid.data.shift());
    } else if (mergedGrid.config.promoteHeaders && !mergedGrid.config.adjustColumnNames) {
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        columnNames = mergedGrid.data.shift()!;
    } else {
        columnNames = Array.from({ length: mergedGrid.data[0].length }, (_, index) => `${defaults.columnName} ${index + 1}`);
    }
    return { columnNames: columnNames, rows: mergedGrid.data };
};

const correctGrid = (grid: MergedGrid): void => {
    if (grid.data.length === 0) {
        // empty grid fix
        grid.config.promoteHeaders = false;
        grid.data.push([""]);
        return;
    }

    const getEmptyArray = (n: number) => Array.from({ length: n }, () => "");
    if (grid.data[0].length === 0) {
        grid.data[0] = [""];
    }
    // replace empty rows
    grid.data.forEach((row, index) => {
        if (row.length === 0) {
            grid.data[index] = getEmptyArray(grid.data[0].length);
        }
    });

    if (grid.config.promoteHeaders && grid.data.length === 1) {
        // table in Excel should have at least 2 rows
        grid.data.push(getEmptyArray(grid.data[0].length));
    }
};

/*
 * Validates the grid, throws an error if the grid is invalid.
 * A valid grid has:
 * - MxN structure.
 * - If promoteHeaders is true - has at least 1 row, and in case adjustColumnNames is false, first row is unique and non empty.
 */
const validateGrid = (grid: MergedGrid): void => {
    validateDataArrayDimensions(grid.data);

    if (grid.config.promoteHeaders && grid.config.adjustColumnNames === false && !validateUniqueAndValidDataArray(grid.data[0])) {
        throw new Error(promotedHeadersCannotBeUsedWithoutAdjustingColumnNamesErr);
    }
};

const validateDataArrayDimensions = (arr: unknown[][]): void => {
    if (arr.length === 0 || arr[0].length === 0) {
        throw new Error(unexpectedErr);
    }

    if (!arr.every((innerArr) => innerArr.length === arr[0].length)) {
        throw new Error(arrayIsntMxNErr);
    }
};

const validateUniqueAndValidDataArray = (arr: string[]): boolean => {
    if (arr.some((element) => element === "")) {
        return false; // Array contains empty elements
    }

    const uniqueSet = new Set(arr);
    return uniqueSet.size === arr.length;
};

const getAdjustedColumnNames = (columnNames: string[] | undefined): string[] => {
    if (columnNames === undefined) {
        throw new Error(unexpectedErr);
    }
    let i = 1;
    // replace empty column names with default names, can still conflict if columns exist, but we handle that later
    columnNames = columnNames.map((columnName) => columnName || `${defaults.columnName} ${i++}`);
    const uniqueNames = new Set<string>();
    return columnNames.map((name) => {
        let uniqueName = name;
        i = 1;
        while (uniqueNames.has(uniqueName)) {
            uniqueName = `${name} (${i++})`;
        }
        uniqueNames.add(uniqueName);
        return uniqueName;
    });
};

export default { parseToTableData };
