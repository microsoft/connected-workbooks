// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { pqUtils, xmlPartsUtils } from "../src/utils";
import { connectionsXmlPath, defaults, section1mPath, textResultType, URLS } from "../src/utils/constants";
import { getPackageComponents } from "../src/utils/mashupDocumentParser";
import { SIMPLE_QUERY_WORKBOOK_TEMPLATE, QUERY_WORKBOOK_TEMPLATE_DIFFRENT_SHEET_NAME } from "../src/workbookTemplate";
import { section1mBlankQueryMock, pqEmptySingleQueryBase64, connectedWorkbookXmlMock, item1Path, item2Path } from "./mocks";
import JSZip from "jszip";
import { describe, test, expect, beforeAll } from '@jest/globals';
import { pqConnectionWithrefreshOnLoadDisable, pqConnectionWithrefreshOnLoadEnable } from "./mocks/xmlMocks";

const getZip = async (template: string) =>
    await JSZip.loadAsync(template, {
        base64: true,
    });

describe("Single query template tests", () => {
    const singleQueryDefaultTemplate = SIMPLE_QUERY_WORKBOOK_TEMPLATE;
    const singleQueryTemplateDiffrenName = QUERY_WORKBOOK_TEMPLATE_DIFFRENT_SHEET_NAME;

    let defaultZipFile;
    let zipFileWithSheetName;

    beforeAll(async () => {
        expect(async () => await getZip(singleQueryDefaultTemplate)).not.toThrow();

        defaultZipFile = await getZip(singleQueryDefaultTemplate);

        expect(async () => await getZip(singleQueryTemplateDiffrenName)).not.toThrow();

        zipFileWithSheetName = await getZip(singleQueryTemplateDiffrenName);
    });

    test("Default template is a valid zip file", async () => {
        expect(defaultZipFile).toBeTruthy();
    });

    test("Template with different name is a valid zip file", async () => {
        expect(zipFileWithSheetName).toBeTruthy();
    });

    test("DataMashup XML exists, and valid PQ Base64 can be extracted", async () => {
        expect(async () => await pqUtils.getDataMashupFile(defaultZipFile)).not.toThrowError();

        const { found, path, value } = await pqUtils.getDataMashupFile(defaultZipFile);

        expect(found).toBeTruthy();
        expect(value).toEqual(pqEmptySingleQueryBase64);
        expect(path).toEqual(item1Path);
    });

    test("ConnectedWorkbook XML exists as item1.xml", async () => {
        const { found, path, xmlString } = await pqUtils.getCustomXmlFile(defaultZipFile, URLS.CONNECTED_WORKBOOK);

        expect(found).toBeTruthy();
        expect(xmlString).toEqual(connectedWorkbookXmlMock);
        expect(path).toEqual(item2Path);
    });

    test("A single blank query named Query1 exists", async () => {
        const base64Str = await pqUtils.getBase64(defaultZipFile);
        const { packageOPC } = getPackageComponents(base64Str!);
        const packageZip = await JSZip.loadAsync(packageOPC);
        const section1m: string | undefined = await packageZip.file(section1mPath)?.async(textResultType);
        if (section1m == undefined) {
            throw new Error("section1m is undefined");
        }
        const hasQuery1 = section1m.includes("Query1");

        expect(hasQuery1).toBeTruthy();
        expect(section1m).toEqual(section1mBlankQueryMock);
    });

    test("A update query with specific sheet name", async () => {
        const queryName = defaults.queryName;
        const sheetName = "SheetNew";
        const refreshOnOpen = false;
        const zipFileWithSheetNameClone = zipFileWithSheetName.clone(); 
        expect(xmlPartsUtils.updateWorkbookSingleQueryAttributes(zipFileWithSheetNameClone, queryName, refreshOnOpen, sheetName)).resolves.toBeUndefined();

        await xmlPartsUtils.updateWorkbookSingleQueryAttributes(zipFileWithSheetNameClone, queryName, refreshOnOpen, sheetName);
        const connectionsXmlString: string | undefined = await zipFileWithSheetNameClone.file(connectionsXmlPath)?.async(textResultType);
        expect(connectionsXmlString).toContain(pqConnectionWithrefreshOnLoadDisable);

    });

    test("A update query with default sheet name", async () => {
        const queryName = defaults.queryName;
        const refreshOnOpen = true;
        const defaultZipFileClone = defaultZipFile.clone(); 
        expect(xmlPartsUtils.updateWorkbookSingleQueryAttributes(defaultZipFileClone, queryName, refreshOnOpen)).resolves.toBeUndefined();

        await xmlPartsUtils.updateWorkbookSingleQueryAttributes(defaultZipFileClone, queryName, refreshOnOpen);
        const connectionsXmlString: string | undefined = await defaultZipFileClone.file(connectionsXmlPath)?.async(textResultType);
        expect(connectionsXmlString).toContain(pqConnectionWithrefreshOnLoadEnable);
    });

    test("sent sheet name does not exist in template", async () => {
        const queryName = defaults.queryName;
        const sheetName = "SheetElse";
        const refreshOnOpen = true;
        expect(xmlPartsUtils.updateWorkbookSingleQueryAttributes(zipFileWithSheetName, queryName, refreshOnOpen, sheetName)).rejects.toEqual(new Error(`Sheet with name ${sheetName} not found`));
    });
});

