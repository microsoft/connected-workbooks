// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TextDecoder, TextEncoder } from "util";
import { replaceSingleQuery, getPackageComponents, editSingleQueryMetadata } from "../src/utils/mashupDocumentParser";
import { arrayUtils, pqUtils } from "../src/utils";
import { section1mNewQueryNameSimpleMock, pqMetadataXmlMockPart1, pqMetadataXmlMockPart2 } from "./mocks";
import JSZip from "jszip";
import { SIMPLE_QUERY_WORKBOOK_TEMPLATE } from "../src/workbookTemplate";
import { section1mPath } from "../src/utils/constants";
import { describe, test, expect } from '@jest/globals';

import util from "util";

(global as any).TextDecoder = TextDecoder;
(global as any).TextEncoder = TextEncoder;

describe("Mashup Document Parser tests", () => {
    test("ReplaceSingleQuery test", async () => {
        const defaultZipFile = await JSZip.loadAsync(SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile);

        if (originalBase64Str) {
            const replacedQueryBase64Str = await replaceSingleQuery(originalBase64Str, "newQueryName", section1mNewQueryNameSimpleMock);
            const buffer = Buffer.from(replacedQueryBase64Str,'base64');
            
            const packageSize = buffer.readInt32LE(4);
            const packageOPC = new Uint8Array(buffer.subarray(8, 8 + packageSize));
            const zip = await JSZip.loadAsync(packageOPC);
            const section1m = await zip.file(section1mPath)?.async("text");
            if (section1m) {
                const mocksection1 = section1mNewQueryNameSimpleMock.replace(/ /g, "");
                expect(section1m.replace(/ /g, "")).toEqual(mocksection1);
            }
        }
    });

    test("Power Query MetadataXml test", async () => {
        const defaultZipFile = await JSZip.loadAsync(SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile);
        if (originalBase64Str) {
            const { metadata } = getPackageComponents(originalBase64Str);
            const newMetadataArray = editSingleQueryMetadata(metadata as Uint8Array, { queryName: "newQueryName" });
            const metadataString = new util.TextDecoder("utf-8").decode(newMetadataArray);
            expect(metadataString.replace(/ /g, "")).toContain(pqMetadataXmlMockPart1.replace(/ /g, ""));
            expect(metadataString.replace(/ /g, "")).toContain(pqMetadataXmlMockPart2.replace(/ /g, ""));
        }
    });
});
