// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { sharedStringsXmlMock, existingSharedStringsXmlMock } from "./mocks";
import { gridUtils, xmlInnerPartsUtils, xmlPartsUtils } from "../src/utils";
import { describe, test, expect } from '@jest/globals';
import JSZip from "jszip";
import { SIMPLE_BLANK_TABLE_TEMPLATE, SIMPLE_QUERY_WORKBOOK_TEMPLATE, WORKBOOK_TEMPLATE_MOVED_TABLE } from "../src/workbookTemplate";

describe("Workbook Manager tests", () => {
    const mockConnectionString = `<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16" mc:Ignorable="xr16">
        <connection id="1" xr16:uid="{86BA784C-6640-4989-A85E-EB4966B9E741}" keepAlive="1" name="Query - Query1" description="Connection to the 'Query1' query in the workbook." type="5" refreshedVersion="7" background="1" saveData="1">
        <dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Query1;" command="SELECT * FROM [Query1]"/></connection></connections>`;

    test("Connection XML attributes contain new query name", async () => {
        const { connectionXmlFileString } = await xmlInnerPartsUtils.updateConnections(mockConnectionString, "newQueryName", true);
        expect(connectionXmlFileString.replace(/ /g, "")).toContain('command="SELECT * FROM [newQueryName]'.replace(/ /g, ""));
        expect(connectionXmlFileString.replace(/ /g, "")).toContain('name="Query - newQueryName"'.replace(/ /g, ""));
        expect(connectionXmlFileString.replace(/ /g, "")).toContain(`description="Connection to the 'newQueryName' query in the workbook."`.replace(/ /g, ""));
    });

    test("Connection XML attributes contain refreshOnLoad value", async () => {
        const { connectionXmlFileString } = await xmlInnerPartsUtils.updateConnections(mockConnectionString, "newQueryName", true);
        expect(connectionXmlFileString.replace(/ /g, "")).toContain('refreshOnLoad="1"');
    });

    test("Connection XML attributes query name with ]", async () => {
        const { connectionXmlFileString } = await xmlInnerPartsUtils.updateConnections(mockConnectionString, "[[name]]]", true);
        expect(connectionXmlFileString.replace(/ /g, "")).toContain("command=\"SELECT*FROM[[[name]]]]]]]\"");
    });

    test("Connection XML attributes query name with no ]", async () => {
        const { connectionXmlFileString } = await xmlInnerPartsUtils.updateConnections(mockConnectionString, "name", true);
        expect(connectionXmlFileString.replace(/ /g, "")).toContain("command=\"SELECT*FROM[name]\"");
    });

    test("Connection XML attributes query name with ] in the middle", async () => {
        const { connectionXmlFileString } = await xmlInnerPartsUtils.updateConnections(mockConnectionString, "[na]me]", true);
        expect(connectionXmlFileString.replace(/ /g, "")).toContain("command=\"SELECT*FROM[[na]]me]]]\"");
    });

    test("SharedStrings XML contains new query name", async () => {
        const { newSharedStrings } = await xmlInnerPartsUtils.updateSharedStrings(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Query1</t></si><si><t/></si></sst>',
            "newQueryName"
        );
        expect(newSharedStrings.replace(/ /g, "")).toContain(sharedStringsXmlMock.replace(/ /g, ""));
    });

    test("Tests SharedStrings update when XML contains query name", async () => {
        const { newSharedStrings } = await xmlInnerPartsUtils.updateSharedStrings(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>newQueryName</t></si><si><t/></si></sst>',
            "newQueryName"
        );
        expect(newSharedStrings.replace(/ /g, "")).toContain(existingSharedStringsXmlMock.replace(/ /g, ""));
    });

    test("SharedStrings XML returns new index", async () => {
        const { sharedStringIndex } = await xmlInnerPartsUtils.updateSharedStrings(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Query1</t></si><si><t/></si></sst>',
            "newQueryName"
        );
        expect(sharedStringIndex).toEqual(2);
    });

    test("SharedStrings XML returns existing index", async () => {
        const { sharedStringIndex } = await xmlInnerPartsUtils.updateSharedStrings(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>newQueryName</t></si><si><t/></si></sst>',
            "newQueryName"
        );
        expect(sharedStringIndex).toEqual(1);
    });

    test("Table XML contains refrshonload value", async () => {
        const { sharedStringIndex, newSharedStrings } = await xmlInnerPartsUtils.updateSharedStrings(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Query1</t></si><si><t/></si></sst>',
            "newQueryName"
        );
        expect(sharedStringIndex).toEqual(2);
        expect(newSharedStrings.replace(/ /g, "")).toContain(sharedStringsXmlMock.replace(/ /g, ""));
    });
    
    test("Table XML contains correct Table reference value with headers included", async () => {
        const singleTableDefaultTemplate = SIMPLE_BLANK_TABLE_TEMPLATE;

        expect(async () => await JSZip.loadAsync(singleTableDefaultTemplate, {
            base64: true,
        })).not.toThrow();

        const defaultZipFile = await JSZip.loadAsync(singleTableDefaultTemplate, {
            base64: true,
        });

        const data = [
            ["ID", "Name", "Income", "Gross", "Bonus"],
            [123, "Alan C", 155000, 155000, 0.15],
            [331, "Tim C", 65000, 13000, 0.12],
            [222, "Bill G", 29501, 8850.3, 0.18],
            [5582, "Mitch M", 87960, 17592, 0.15],
            [43, "Dan F", 197296, 19729.6, 0.22],
            [22, "Perry T-P", 186006, 37201.2, 0.4],
            [335, "Mdrake", 197136, 78854.4, 0.1],
            [6590, "Dr P", 139636, 41890.8, 0.13],
        ];
        const tableData = gridUtils.parseToTableData({ data: data, config: { promoteHeaders: true, adjustColumnNames: true } });

        await xmlPartsUtils.updateWorkbookDataAndConfigurations(defaultZipFile, undefined, tableData);
        expect(await defaultZipFile.file("xl/tables/table1.xml")?.async("text")).toContain("A1:E9");

    });

    test("Table XML contains correct Table reference value without headers included", async () => {
        const singleTableDefaultTemplate = SIMPLE_BLANK_TABLE_TEMPLATE;

        expect(async () => await JSZip.loadAsync(singleTableDefaultTemplate, {
            base64: true,
        })).not.toThrow();

        const defaultZipFile = await JSZip.loadAsync(singleTableDefaultTemplate, {
            base64: true,
        });

        const data = [
            ["ID", "Name", "Income", "Gross", "Bonus"],
            [123, "Alan C", 155000, 155000, 0.15],
            [331, "Tim C", 65000, 13000, 0.12],
            [222, "Bill G", 29501, 8850.3, 0.18],
            [5582, "Mitch M", 87960, 17592, 0.15],
            [43, "Dan F", 197296, 19729.6, 0.22],
            [22, "Perry T-P", 186006, 37201.2, 0.4],
            [335, "Mdrake", 197136, 78854.4, 0.1],
            [6590, "Dr P", 139636, 41890.8, 0.13],
        ];
        const tableData = gridUtils.parseToTableData({ data: data, config: { promoteHeaders: false, adjustColumnNames: true } });

        await xmlPartsUtils.updateWorkbookDataAndConfigurations(defaultZipFile, undefined, tableData);
        expect(await defaultZipFile.file("xl/tables/table1.xml")?.async("text")).toContain("A1:E10");
    });

    test("Table XML contains correct Table reference value using template", async () => {
        const movedTableDefaultTemplate = WORKBOOK_TEMPLATE_MOVED_TABLE;
        expect(async () => await JSZip.loadAsync(movedTableDefaultTemplate, {
            base64: true,
        })).not.toThrow();

        const templateMovedZipFile :any = await JSZip.loadAsync(movedTableDefaultTemplate, {
            base64: true,
        });
        const data = [
            ["ID", "Name", "Income", "Gross", "Bonus"],
            [123, "Alan C", 155000, 155000, 0.15],
            [331, "Tim C", 65000, 13000, 0.12],
            [222, "Bill G", 29501, 8850.3, 0.18],
            [5582, "Mitch M", 87960, 17592, 0.15],
            [43, "Dan F", 197296, 19729.6, 0.22],
            [22, "Perry T-P", 186006, 37201.2, 0.4],
            [335, "Mdrake", 197136, 78854.4, 0.1],
            [6590, "Dr P", 139636, 41890.8, 0.13],
        ];
        const tableData = gridUtils.parseToTableData({ data: data, config: { promoteHeaders: true, adjustColumnNames: true } });

        await xmlPartsUtils.updateWorkbookDataAndConfigurations(templateMovedZipFile, { templateFile: templateMovedZipFile }, tableData);
        expect(await templateMovedZipFile.file("xl/tables/table1.xml")?.async("text")).toContain("B2:F10");
    });
});

