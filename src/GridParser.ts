import { defaults, gridNotFoundErr } from "./utils/constants";
import { Grid, TableData } from "./types";
import { tableUtils } from "./utils";

export const parseToTableData = async (grid: Grid): Promise<TableData | undefined> => {
    if (!grid) {
        return undefined;
    }

    const columnNames: string[] = await generateColumnNames(grid);
    const rows: string[][] = parseGridRows(grid);

    return { columnNames: columnNames, rows: rows };
};

const parseGridRows = (grid: Grid): string[][] => {
    const gridData: (string | number | boolean)[][] = grid.data;
    if (!gridData) {
        throw new Error(gridNotFoundErr);
    }

    const rows: string[][] = [];
    if (!grid.headerCofings?.promoteHeaders) {
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

const generateColumnNames = async (grid: Grid): Promise<string[]> => {
    const columnNames: string[] = [];
    if (!grid.headerCofings?.promoteHeaders) {
        for (let i = 0; i < grid.data[0].length; i++) {
            columnNames.push(`${defaults.columnName} ${i + 1}`);
        }

        return columnNames;
    }

    if (grid.headerCofings?.promoteHeaders) {
        return await tableUtils.getColumnNames(grid.data[0]);
    }

    return columnNames;
};
