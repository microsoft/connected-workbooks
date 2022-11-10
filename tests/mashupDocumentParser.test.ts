import { TextDecoder, TextEncoder } from 'util';
import MashupHandler from "../src/mashupDocumentParser";
import { arrayUtils, pqUtils } from "../src/utils";
import { simpleQueryMock, section1mSimpleQueryMock, section1mNewQueryNameSimpleMock, section1mNewQueryNameBlankMock, relationshipInfo } from "./mocks";
import base64 from "base64-js";
import JSZip from "jszip";
import WorkbookTemplate from "../src/workbookTemplate";
import { section1mPath } from "../src/constants";
import { Metadata } from "../src/types";

(global as any).TextDecoder = TextDecoder;
(global as any).TextEncoder = TextEncoder;

describe("Mashup Document Parser tests", () => {
    test("ReplaceSingleQuery test", async () => {
        const mashupHandler = new MashupHandler();

        const defaultZipFile = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile);

          if (originalBase64Str) {
            const replacedQueryBase64Str = await mashupHandler.ReplaceSingleQuery(originalBase64Str, "newQueryName", simpleQueryMock);
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

    test("queryMetadataEntries test", async () => {
        const defaultZipFile = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile); 
        if  (originalBase64Str) {
            const handler = new MashupHandler() as any;
            const { metadata } = handler.getPackageComponents(originalBase64Str);
            const newMetadataArray = handler.editSingleQueryMetadata(metadata, {queryName: "newQueryName"});
            const mashupArray = new arrayUtils.ArrayReader(newMetadataArray.buffer);
            const metadataVersion = mashupArray.getBytes(4);
            const metadataXmlSize = mashupArray.getInt32();
            const metadataXml = mashupArray.getBytes(metadataXmlSize);

            //parse metdataXml
            const util= require('util');
            const metadataString =new util.TextDecoder("utf-8").decode(metadataXml);
            const parser = new DOMParser();
            const parsedMetadata = parser.parseFromString(metadataString, "text/xml");
            const entries = parsedMetadata.getElementsByTagName("Entry");
            if (entries && entries.length) {
                for (let i = 0; i < entries.length; i++) {
                    const entry = entries[i];
                    const entryAttributesArr = [...entry.attributes]; 
                    const entryProp = entryAttributesArr.find((prop) => {
                        return prop?.name === "Type"});
                    if (entryProp?.nodeValue == "RelationshipInfoContainer") {
                             expect(entry.getAttribute("Value")).toContain("Section1/newQueryName");
                    }

                    if (entryProp?.nodeValue == "ResultType") {
                        expect(entry.getAttribute("Value")).toEqual("sTable");
                    }
                    if (entryProp?.nodeValue == "FillTarget") {
                        expect(entry.getAttribute("Value")).toEqual("snewQueryName");
                    }
                    if (entryProp?.nodeValue == "FillColumnNames") {
                        expect(entry.getAttribute("Value")).toContain("newQueryName");
                    }

                    if (entryProp?.nodeValue == "FillLastUpdated") {
                        const nowTime = new Date().toISOString();
                        expect(entry.getAttribute("Value")).toContain(nowTime.substring(0, nowTime.length - 10));
                    }
                }
            }


        }
    })
});
