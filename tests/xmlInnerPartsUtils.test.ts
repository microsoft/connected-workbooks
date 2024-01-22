// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { sharedStringsXmlMock, existingSharedStringsXmlMock } from "./mocks";
import { xmlInnerPartsUtils } from "../src/utils";

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

    test("Connections XML contains new connection", async () => {
        const serializer = new XMLSerializer();
        const mockXml = new DOMParser().parseFromString(mockConnectionString, "text/xml");
        const newConnectionsXml: Document = await xmlInnerPartsUtils.addNewConnection(mockXml, "newQueryName");
        const newConnectionsXmlString = serializer.serializeToString(newConnectionsXml);
        expect((newConnectionsXmlString.match(/<connection id/g) || []).length).toEqual(2);

    });
});
