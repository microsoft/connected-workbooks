// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { defaults } from "../utils/constants";
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
    let columnNames: string[] = [];
    if (mergedGrid.config.promoteHeaders && mergedGrid.config.adjustColumnNames) {
        columnNames = getAdjustedColumnNames(mergedGrid.data.shift());
    } else if (mergedGrid.config.promoteHeaders && !mergedGrid.config.adjustColumnNames) {
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        columnNames = mergedGrid.data.shift()!;
    } else {
        columnNames = Array.from({ length: grid.data[0].length }, (_, index) => `${defaults.columnName} ${index + 1}`);
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

const getAdjustedColumnNames = (columnNames: string[] | undefined): string[] => {
    if (columnNames === undefined) {
        throw new Error("Unexpected");
    }
    columnNames = columnNames.map((columnName) => columnName || defaults.columnName);
    const uniqueNames = new Set<string>();
    return columnNames.map((name) => {
        let uniqueName = name;
        let index = 1;
        while (uniqueNames.has(uniqueName)) {
            uniqueName = `${name} (${index++})`;
        }
        uniqueNames.add(uniqueName);
        return uniqueName;
    });
};

export default { parseToTableData };
