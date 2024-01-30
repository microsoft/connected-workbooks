// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DataTypes } from "../src/types";
import { dateTimeUtils, documentUtils } from "../src/utils";
import { element } from "../src/utils/constants";

describe("Document Utils tests", () => {
    test("ResolveType date not supported success", () => {
        expect(documentUtils.resolveType("10:00:59 PM", false, dateTimeUtils.dateTimeFormatArr[3])).toEqual(DataTypes.dateTime);
        expect(documentUtils.resolveType("1/1/2024", false, dateTimeUtils.dateTimeFormatArr[0])).toEqual(DataTypes.dateTime);
        expect(documentUtils.resolveType("10:00 AM", false, dateTimeUtils.dateTimeFormatArr[2])).toEqual(DataTypes.dateTime);
    });

    test("ResolveType unsupported datetime format", () => {
        expect(documentUtils.resolveType("5-4-2023 00:00", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("13/1/2024", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("1/32/2024", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("1/31  /2024", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("30:00", false, undefined)).toEqual(DataTypes.string);
    });

    test("ResolveType string success", () => {
        expect(documentUtils.resolveType("sTrIng", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("True", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("False", false, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("    ", false, undefined)).toEqual(DataTypes.string);
    });

    test("ResolveType boolean success", () => {
        expect(documentUtils.resolveType("true", false, undefined)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("   true", false, undefined)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("false", false, undefined)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("   false", false, undefined)).toEqual(DataTypes.boolean);
    });

    test("ResolveType number success", () => {
        expect(documentUtils.resolveType("100000", false, undefined)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.00", false, undefined)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50", false, undefined)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50", false, undefined)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23450", false, undefined)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23.4.50", false, undefined)).toEqual(DataTypes.string);
    });

    test("ResolveType header row success", () => {
        expect(documentUtils.resolveType("100000", true, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("true", true, undefined)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("string", true, undefined)).toEqual(DataTypes.string);
    });

    test("Cell Data Element preserves spaces", () => {
        const doc = document.implementation.createDocument("", "", null);
        const cell: Element = doc.createElementNS("", element.kindCell);
        const cellData: Element = doc.createElementNS("", element.cellValue);
        documentUtils.updateCellData("     ", cell, cellData, false, []);
        expect(cellData.getAttribute("xml:space")).toEqual("preserve");
        cellData.removeAttribute("xml:space");
        documentUtils.updateCellData("a     ", cell, cellData, false, []);
        expect(cellData.getAttribute("xml:space")).toEqual("preserve");
        cellData.removeAttribute("xml:space");
        documentUtils.updateCellData("     a", cell, cellData, false, []);
        expect(cellData.getAttribute("xml:space")).toEqual("preserve");
        cellData.removeAttribute("xml:space");
        documentUtils.updateCellData("a     a", cell, cellData, false, []);
        expect(cellData.getAttribute("xml:space")).toBeNull();
    });
});
