import { gridNotFoundErr, invalidColumnNamesErr } from "./utils/constants";
import { Grid, TableData } from "./types";
import tableUtils from "./utils/tableUtils";

export const parseToTableData = (grid: Grid): TableData | undefined => {
    if (!grid) {
        return undefined;
    }

    const columnNames: string[] = generateColumnNames(grid);
    if (tableUtils.validateColumnNames(columnNames) === false) {
        throw new Error(invalidColumnNamesErr);
    }

    const rows: string[][] = parseGridRows(grid);

    return { columnNames: columnNames, rows: rows };
};

const parseGridRows = (grid: Grid): string[][] => {
    const gridData: (string | number | boolean)[][] = grid.data;
    if (!gridData) {
        throw new Error(gridNotFoundErr);
    }

    const rows: string[][] = [];
    if (!grid.promoteHeaders) {
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
            row.push(cellValue.toString());
        }

        rows.push(row);
    }

    return rows;
};

const generateColumnNames = (grid: Grid): string[] => {
    if (grid.promoteHeaders) {
        const columnNames: string[] = grid.data[0].map((columnName) => columnName.toString());
        if (!tableUtils.validateColumnNames(columnNames)) {
            throw new Error(invalidColumnNamesErr);
        }
        
        return columnNames;
    }

    const columnNames: string[] = [];
    for (let i = 0; i < grid.data[0].length; i++) {
        columnNames.push(`Column ${i + 1}`);
    }

    return columnNames;
};
