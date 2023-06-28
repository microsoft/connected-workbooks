export const extractTableValues = (table: HTMLTableElement): string[][] => {
    const rows: string[][] = [];

    // Extract values from each row
    for (let i = 0; i < table.rows.length; i++) {
        const row = table.rows[i];
        const rowData: string[] = [];

        for (let j = 0; j < row.cells.length; j++) {
            const cell = row.cells[j];
            rowData.push(cell.textContent || "");
        }

        rows.push(rowData);
    }

    return rows;
};
