import { GridNotFoundErr, headerNotFoundErr, invalidDataTypeErr, invalidValueInColumnErr } from "./constants";
import { ColumnMetadata, dataTypes, Grid, TableData } from "./types";

export default class GridParser {
    public parseToTableData(initialDataGrid: Grid): TableData | undefined {
        if (!initialDataGrid) {
            return undefined;
        }

        this.validateGridHeader(initialDataGrid);
        const data: string[][] = this.parseGridData(initialDataGrid, initialDataGrid.Header);
        
        return { columnMetadata : initialDataGrid.Header, data: data };
    }

    private parseGridData(initialDataGrid: Grid, columnMetadata: ColumnMetadata[]) {
        const gridData: (string | number | boolean)[][] = initialDataGrid.GridData;
        if (!gridData) {
            throw new Error(GridNotFoundErr);
        }
        
        const tableData: string[][] = [];
        for (const rowData of gridData) {
            const row: string[] = [];
            var colIndex: number = 0;
            for (const prop in rowData) {
                const dataType: dataTypes = columnMetadata[colIndex].type;
                const cellValue: string | number | boolean = rowData[prop];
                if (dataType == dataTypes.number) {
                    if (isNaN(Number(cellValue))) {
                        throw new Error(invalidValueInColumnErr);
                    }
                }

                if (dataType == dataTypes.boolean) {
                    if (cellValue != "1" && cellValue != "0") {
                        throw new Error(invalidValueInColumnErr);
                    }
                }

                row.push(rowData[prop].toString());
                colIndex++;
            }
            tableData.push(row);
        }

        return tableData;
    }

    private validateGridHeader(data: Grid) {
        const headerData: ColumnMetadata[] = data.Header;
        if (!headerData) {
            throw new Error(headerNotFoundErr);
        }

        for (const prop in headerData) {
            if (!(headerData[prop].type in dataTypes)) { 
                throw new Error(invalidDataTypeErr);
            }
        }
    }
}
