import { gridNotFoundErr } from "./utils/constants";
import { Grid, TableData, TableDataParser } from "./types";

export default class GridParser implements TableDataParser {
    public parseToTableData(initialDataGrid: Grid): TableData | undefined {
        if (!initialDataGrid) {
            return undefined;
        }

        const columnNames: string[] = this.generateColumnNames(initialDataGrid);
        const rows: string[][] = this.parseGridRows(initialDataGrid);

        return { columnNames: columnNames, rows: rows };
    }

    private parseGridRows(initialDataGrid: Grid): string[][] {
        const gridData: (string | number | boolean)[][] = initialDataGrid.gridData;
        if (!gridData) {
            throw new Error(gridNotFoundErr);
        }

        const rows: string[][] = [];
        for (const rowData of gridData) {
            const row: string[] = [];
            let colIndex = 0;
            for (const prop in rowData) {
                const cellValue: string | number | boolean = rowData[prop];
                row.push(cellValue.toString());
                colIndex++;
            }
            
            rows.push(row);
        }

        return rows;
    }

    private generateColumnNames(initialDataGrid: Grid): string[] {  
        if (!initialDataGrid.promoteHeaders) {
            return initialDataGrid.gridData[0].map((columnName) => (columnName.toString()));
        }
        
        let columnNames: string[] = [];
        for (let i = 0; i < initialDataGrid.gridData[0].length; i++) {
            columnNames.push(`Column ${i + 1}`);
        }
        
        return columnNames;
    }

}
