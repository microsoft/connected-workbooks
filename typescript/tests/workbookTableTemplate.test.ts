// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SIMPLE_BLANK_TABLE_TEMPLATE } from "../src/workbookTemplate";
import { pqUtils } from "../src/utils";
import { queryTableXmlPath, tableXmlPath, textResultType } from "../src/utils/constants";
import JSZip from "jszip";
import { describe, test, expect,beforeAll } from '@jest/globals';

const getZip = async (template: string) =>
    await JSZip.loadAsync(template, {
        base64: true,
    });

describe("Single blank table template tests", () => {
    const simpleTableDefaultTemplate = SIMPLE_BLANK_TABLE_TEMPLATE;
    let defaultZipFile;

    beforeAll(async () => {
        expect(async () => await getZip(simpleTableDefaultTemplate)).not.toThrow();

        defaultZipFile = await getZip(simpleTableDefaultTemplate);
    });

    test("Default template is a valid zip file", () => {
        expect(defaultZipFile).toBeTruthy();
    });

    test("DataMashup XML doesn't exists", async () => {
        try {
            await pqUtils.getDataMashupFile(defaultZipFile);
        } catch (error) {
            expect(error).toBeTruthy();
        }
    });

    test("A single blank table exists", async () => {
        const tableXml: string | undefined = await defaultZipFile.file(tableXmlPath)?.async(textResultType);
        expect(tableXml).toContain('name="Table1" displayName="Table1" ref="A1:A2"');
    });

    test("Does not contains query table", async () => {
        const queryTableXml: string | undefined = await defaultZipFile.file(queryTableXmlPath)?.async(textResultType);
        expect(queryTableXml).toBeFalsy();
    });
});
