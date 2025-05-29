// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { htmlUtils } from "../src/utils";
import { JSDOM } from "jsdom";
import { describe, test, expect } from '@jest/globals';

// Create a JSDOM instance
const { window } = new JSDOM("<!DOCTYPE html><html><body></body></html>");
const document = window.document;

// Helper function to create table rows and cells
const createRowWithCells = (cellValues: string[], celltypes = "td") => {
    const row = document.createElement("tr");
    cellValues.forEach((value) => {
        const cell = document.createElement(celltypes);
        cell.textContent = value;
        row.appendChild(cell);
    });
    return row;
};

describe("extractTableValues", () => {
    test("returns an empty array for an empty table", () => {
        // Create an empty table element
        const table = document.createElement("table");

        // Call the method and expect an empty array as the result
        expect(htmlUtils.extractTableValues(table)).toEqual([]);
    });

    test("returns an empty row for an empty tr", () => {
        const table = document.createElement("table");
        const tbody = document.createElement("tbody");

        const row = createRowWithCells([]);

        tbody.appendChild(row);
        table.appendChild(tbody);

        expect(htmlUtils.extractTableValues(table)).toEqual([[]]);
    });

    test("extracts values correctly from a table with multiple rows and cells", () => {
        const table = document.createElement("table");
        const tbody = document.createElement("tbody");

        const row1 = createRowWithCells(["A", "B"]);
        const row2 = createRowWithCells(["C", "D"]);

        tbody.appendChild(row1);
        tbody.appendChild(row2);

        table.appendChild(tbody);

        const expectedResult = [
            ["A", "B"],
            ["C", "D"],
        ];

        expect(htmlUtils.extractTableValues(table)).toEqual(expectedResult);
    });

    test("handles empty cells by using an empty string", () => {
        const table = document.createElement("table");
        const tbody = document.createElement("tbody");

        const row1 = createRowWithCells(["A", ""]);
        const row2 = createRowWithCells(["", "D"]);

        tbody.appendChild(row1);
        tbody.appendChild(row2);

        table.appendChild(tbody);

        const expectedResult = [
            ["A", ""],
            ["", "D"],
        ];

        expect(htmlUtils.extractTableValues(table)).toEqual(expectedResult);
    });

    test("handle table header (th) cells", () => {
        const table = document.createElement("table");
        const tbody = document.createElement("tbody");

        const headerRow = createRowWithCells(["Header 1", "Header 2"], "th");
        const dataRow = createRowWithCells(["A", "B"]);

        tbody.appendChild(headerRow);
        tbody.appendChild(dataRow);

        table.appendChild(tbody);

        const expectedResult = [
            ["Header 1", "Header 2"],
            ["A", "B"],
        ];

        expect(htmlUtils.extractTableValues(table)).toEqual(expectedResult);
    });

    test("handles tables with multiple tbody elements", () => {
        const table = document.createElement("table");

        const tbody1 = document.createElement("tbody");
        const row1 = createRowWithCells(["A", "B"]);
        tbody1.appendChild(row1);
        table.appendChild(tbody1);

        const tbody2 = document.createElement("tbody");
        const row2 = createRowWithCells(["C", "D"]);
        tbody2.appendChild(row2);
        table.appendChild(tbody2);

        const expectedResult = [
            ["A", "B"],
            ["C", "D"],
        ];
        expect(htmlUtils.extractTableValues(table)).toEqual(expectedResult);
    });

    test("handles tables that are not MxN", () => {
        const table = document.createElement("table");

        const tbody1 = document.createElement("tbody");
        const row1 = createRowWithCells(["A", "B"]);
        tbody1.appendChild(row1);
        table.appendChild(tbody1);

        const tbody2 = document.createElement("tbody");
        const row2 = createRowWithCells(["C", "D", "E"]);
        tbody2.appendChild(row2);
        table.appendChild(tbody2);

        const expectedResult = [
            ["A", "B"],
            ["C", "D", "E"],
        ];
        expect(htmlUtils.extractTableValues(table)).toEqual(expectedResult);
    });
});
