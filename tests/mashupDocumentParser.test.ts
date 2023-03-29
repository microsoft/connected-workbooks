import { TextDecoder, TextEncoder } from 'util';
import MashupHandler from "../src/mashupDocumentParser";
import { arrayUtils, pqUtils } from "../src/utils";
import { simpleQueryMock, section1mNewQueryNameSimpleMock, pqMetadataXmlMockPart1, pqMetadataXmlMockPart2 } from "./mocks";
import base64 from "base64-js";
import JSZip from "jszip";
import WorkbookTemplate from "../src/workbookTemplate";
import { section1mPath } from "../src/constants";

(global as any).TextDecoder = TextDecoder;
(global as any).TextEncoder = TextEncoder;

describe("Mashup Document Parser tests", () => {
    test("ReplaceSingleQuery test", async () => {
        const mashupHandler = new MashupHandler();

        const defaultZipFile = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile);

          if (originalBase64Str) {
            const replacedQueryBase64Str = await mashupHandler.ReplaceSingleQuery(originalBase64Str, "Query1", section1mNewQueryNameSimpleMock);
            const buffer = base64.toByteArray(replacedQueryBase64Str).buffer;
            const mashupArray = new arrayUtils.ArrayReader(buffer);
            const startArray = mashupArray.getBytes(4);
            const packageSize = mashupArray.getInt32();
            const packageOPC = mashupArray.getBytes(packageSize);
            const zip = await JSZip.loadAsync(packageOPC);
            const section1m = await zip.file(section1mPath)?.async("text");
            if (section1m) {
                const mocksection1 = section1mNewQueryNameSimpleMock.replace(/ /g, "");
                expect(section1m.replace(/ /g, "")).toEqual(mocksection1);
            }
          }
    })

    test("Power Query MetadataXml test", async () => {
        const defaultZipFile = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile); 
        if  (originalBase64Str) {
            const handler = new MashupHandler() as any;
            const { metadata } = handler.getPackageComponents(originalBase64Str);
            const newMetadataArray = handler.editSingleQueryMetadata(metadata, {queryName: "newQueryName"});
            const util = require('util');
            const metadataString = (new util.TextDecoder("utf-8").decode(newMetadataArray));
            expect(metadataString.replace(/ /g, "")).toContain(pqMetadataXmlMockPart1.replace(/ /g, ""));
            expect(metadataString.replace(/ /g, "")).toContain(pqMetadataXmlMockPart2.replace(/ /g, ""));
        }
    })
});
