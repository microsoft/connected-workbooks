export const extractTableValues = (table: HTMLTableElement): [string[], string[][]] => {
    const headers: string[] = [];
    const rows: string[][] = [];

    // Extract headers
    const headerRow = table.rows[0];
    for (let i = 0; i < headerRow.cells.length; i++) {
        const cell = headerRow.cells[i];
        headers.push(cell.textContent || "");
    }

    // Extract values from each row
    for (let i = 1; i < table.rows.length; i++) {
        const row = table.rows[i];
        const rowData: string[] = [];

        for (let j = 0; j < row.cells.length; j++) {
            const cell = row.cells[j];
            rowData.push(cell.textContent || "");
        }

        rows.push(rowData);
    }

    return [headers, rows];
};
