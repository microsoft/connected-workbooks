// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DataTypes } from "../src/types";
import { documentUtils } from "../src/utils";
import { element } from "../src/utils/constants";

describe("Document Utils tests", () => {
    test("ResolveType date not supported success", () => {
        expect(documentUtils.resolveType("5-4-2023 00:00", false)).toEqual(DataTypes.string);
    });

    test("ResolveType string success", () => {
        expect(documentUtils.resolveType("sTrIng", false)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("True", false)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("False", false)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("    ", false)).toEqual(DataTypes.string);
    });

    test("ResolveType boolean success", () => {
        expect(documentUtils.resolveType("true", false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("   true", false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("false", false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("   false", false)).toEqual(DataTypes.boolean);
    });

    test("ResolveType number success", () => {
        expect(documentUtils.resolveType("100000", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.00", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23450", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23.4.50", false)).toEqual(DataTypes.string);
    });

    test("ResolveType header row success", () => {
        expect(documentUtils.resolveType("100000", true)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("true", true)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("string", true)).toEqual(DataTypes.string);
    });

    test("Cell Data Element preserves spaces", () => {
        const doc = document.implementation.createDocument("", "", null);
        const cell: Element = doc.createElementNS("", element.kindCell);
        const cellData: Element = doc.createElementNS("", element.cellValue);
        documentUtils.updateCellData("     ", cell, cellData, false);
        expect(cellData.getAttribute("xml:space")).toEqual("preserve");
        cellData.removeAttribute("xml:space");
        documentUtils.updateCellData("a     ", cell, cellData, false);
        expect(cellData.getAttribute("xml:space")).toEqual("preserve");
        cellData.removeAttribute("xml:space");
        documentUtils.updateCellData("     a", cell, cellData, false);
        expect(cellData.getAttribute("xml:space")).toEqual("preserve");
        cellData.removeAttribute("xml:space");
        documentUtils.updateCellData("a     a", cell, cellData, false);
        expect(cellData.getAttribute("xml:space")).toBeNull();
    });

    test("Test convert column number To Excel Column", () => {
        expect(documentUtils.convertToExcelColumn(0)).toEqual("A");
        expect(documentUtils.convertToExcelColumn(701)).toEqual("ZZ");
        expect(documentUtils.convertToExcelColumn(16383)).toEqual("XFD");
        try {
            documentUtils.convertToExcelColumn(16384);
        } catch (e) {
            expect(e.message).toEqual("Column index out of range");
        }
    });
});
