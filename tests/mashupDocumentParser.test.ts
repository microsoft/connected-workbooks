import { TextDecoder, TextEncoder } from 'util';
import MashupHandler from "../src/mashupDocumentParser";
import { arrayUtils, pqUtils } from "../src/utils";
import { simpleQueryMock, section1mNewQueryNameSimpleMock, pqMetadataXmlMockPart1, pqMetadataXmlMockPart2, pqMetadataXmlMock } from "./mocks";
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


    test("Create new connection only stableEntries metadata item", async () => {
        const defaultZipFile = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile); 
        if  (originalBase64Str) {
            const handler = new MashupHandler() as any;
            const { metadata } = handler.getPackageComponents(originalBase64Str);
            const newMetadataArray = handler.editSingleQueryMetadata(metadata, {queryName: "Query1"});
            const util = require('util');
            const metadataString = (new util.TextDecoder("utf-8").decode(newMetadataArray));
            const metadataDoc = new DOMParser().parseFromString(metadataString, "text/xml");  
            const stableEntriesItem = handler.createStableEntriesItem(metadataDoc, "Query2");
            compareConnectionOnlyStableEntries(stableEntriesItem);
            compareConnectionOnlyItemLocation(stableEntriesItem, "Query2");
        }
    })

    test("Create new connection only src metadata item", async () => {
        const defaultZipFile = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile); 
        if  (originalBase64Str) {
            const handler = new MashupHandler() as any;
            const { metadata } = handler.getPackageComponents(originalBase64Str);
            const util = require('util');
            const metadataString = (new util.TextDecoder("utf-8").decode(metadata));
            const metadataDoc = new DOMParser().parseFromString(metadataString, "text/xml");  
            const sourceItem = handler.createSourceItem(metadataDoc, "Query2");
            compareConnectionOnlySrcItemLocation(sourceItem, "Query2");
        }
    })

    test("Add connection only query to metadata", async () => {
        const handler = new MashupHandler() as any;
        const connectionOnlyMetadataStr = await handler.updateConnectionOnlyMetadataStr(pqMetadataXmlMock, "Query2");
        expect((connectionOnlyMetadataStr.match(new RegExp("<Item>", "g")) || []).length).toEqual(5);
        
    })
    
});

function compareConnectionOnlyItemLocation(stableEntriesItem: Element, queryName: string) {
    const path = [...stableEntriesItem.getElementsByTagName("ItemPath")][0];
    expect(path.textContent).toEqual(`Section1/${queryName}`);
    const type = [...stableEntriesItem.getElementsByTagName("ItemType")][0];
    expect(type.textContent).toEqual("Formula");
}

function compareConnectionOnlySrcItemLocation(stableEntriesItem: Element, queryName: string) {
    const path = [...stableEntriesItem.getElementsByTagName("ItemPath")][0];
    expect(path.textContent).toEqual(`Section1/${queryName}/Source`);
    const type = [...stableEntriesItem.getElementsByTagName("ItemType")][0];
    expect(type.textContent).toEqual("Formula");
}

function compareConnectionOnlyStableEntries(item: Element) {
    const entries = [...item.getElementsByTagName("Entry")];
    expect(entries.length).toEqual(6);
    entries.forEach(entry => {
        if (entry.getAttribute("type") === "IsPrivate") {
            expect(entry.getAttribute("value")).toEqual("l0");
            }
        if (entry.getAttribute("type") === "FillEnabled") {
            expect(entry.getAttribute("value")).toEqual("l0");
            }
        if (entry.getAttribute("type") === "FillObjectType") {
            expect(entry.getAttribute("value")).toEqual("sConnectionOnly");
            }
        if (entry.getAttribute("type") === "FillToDataModelEnabled") {
            expect(entry.getAttribute("value")).toEqual("l0");
            }
        if (entry.getAttribute("type") === "ResultType") {
            expect(entry.getAttribute("value")).toEqual("sTable");
            }
    });
}

