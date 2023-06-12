import { gridNotFoundErr, headerNotFoundErr, invalidDataTypeErr, invalidValueInColumnErr } from "./constants";
import { ColumnMetadata, DataTypes, Grid, TableData, TableDataParser} from "./types";

export default class GridParser implements TableDataParser {
    public parseToTableData(initialDataGrid: Grid): TableData | undefined {
        if (!initialDataGrid) {
            return undefined;
        }

        this.validateGridHeader(initialDataGrid);
        const rows: string[][] = this.parseGridRows(initialDataGrid, initialDataGrid.header);

        return { columnMetadata: initialDataGrid.header, rows: rows };
    }

    private parseGridRows(initialDataGrid: Grid, columnMetadata: ColumnMetadata[]) : string[][] {
        const gridData: (string | number | boolean)[][] = initialDataGrid.gridData;
        if (!gridData) {
            throw new Error(gridNotFoundErr);
        }
        
        const rows: string[][] = [];
        for (const rowData of gridData) {
            const row: string[] = [];
            var colIndex: number = 0;
            for (const prop in rowData) {
                const dataType: DataTypes = columnMetadata[colIndex].type;
                const cellValue: string | number | boolean = rowData[prop];
                if (dataType == DataTypes.number) {
                    if (isNaN(Number(cellValue))) {
                        throw new Error(invalidValueInColumnErr);
                    }
                }
                else if (dataType == DataTypes.boolean) {
                    if (cellValue != "1" && cellValue != "0") {
                        throw new Error(invalidValueInColumnErr);
                    }
                }

                row.push(rowData[prop].toString());
                colIndex++;
            }
            rows.push(row);
        }

        return rows;
    }

    private validateGridHeader(data: Grid) {
        const headerData: ColumnMetadata[] = data.header;
        if (!headerData) {
            throw new Error(headerNotFoundErr);
        }
        
        for (const prop in headerData) {
            if (!(headerData[prop].type in DataTypes)) { 
                throw new Error(invalidDataTypeErr);
            }
        }
    }
}

