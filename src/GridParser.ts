import { dateFormats, GridNotFoundErr, headerNotFoundErr, invalidDataTypeErr, invalidFormatTypeErr, invalidMissingFormatFromDateTimeErr, invalidValueInColumnErr, dateFormatsRegex, milliSecPerDay, numberOfDaysTillExcelBeginYear, monthsbeforeLeap } from "./constants";
import { ColumnMetadata, dataTypes, Grid, TableData } from "./types";

export class GridParser {
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

                if (dataType == dataTypes.dateTime) {
                    if (dateFormatsRegex[columnMetadata[colIndex].format!].test(cellValue.toString())) {
                        throw new Error(invalidValueInColumnErr); 
                    }
                    rowData[prop] = this.convertToExcelDate(cellValue.toString());
                }

                row.push(rowData[prop].toString());
                colIndex++;
            }
            tableData.push(row);
        }

        return tableData;
    }

    private convertToExcelDate(dateStr: string) {
        const [month, day, year, hour, minute] = dateStr.split(/[\/: ]/);
        const date = new Date(Date.UTC(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hour), parseInt(minute), 0,0)).getTime();
        // Excel incorrectly assumes that the year 1900 is a leap year. This is a workaround for that
        if (parseInt(year) == 1900 && parseInt(month) <= monthsbeforeLeap) {
            return ((date + numberOfDaysTillExcelBeginYear*milliSecPerDay) / (milliSecPerDay)) - 1;
        }
        return (date + numberOfDaysTillExcelBeginYear*milliSecPerDay) / (milliSecPerDay);
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

            if (headerData[prop].type == dataTypes.dateTime && headerData[prop].format == undefined) {
                throw new Error(invalidMissingFormatFromDateTimeErr);
            }

            if (headerData[prop].format != undefined && !(headerData[prop].format! in dateFormats)) {
                throw new Error(invalidFormatTypeErr);
            }

        }
    }
}

