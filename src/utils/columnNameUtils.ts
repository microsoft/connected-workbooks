// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { InvalidColumnNameErr, defaults } from "../utils/constants";

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

export default { getNextAvailableColumnName, getAdjustedColumnNames, getRawColumnNames };
