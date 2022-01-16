import workbookTemplate from "../src/workbookTemplate";
import { pqUtils } from "../src/utils";
import { URLS } from "../src/constants";
import MashupHandler from "../src/mashupDocumentParser";
import {
    section1mBlankQueryMock,
    pqEmptySingleQueryBase64,
    connectedWorkbookXmlMock,
    item1Path,
    item2Path,
} from "./mocks";
import JSZip from "jszip";

const getZip = async (template: string) =>
    await JSZip.loadAsync(template, {
        base64: true,
    });

describe("Single query template tests", () => {
    const singleQueryDefaultTemplate =
        workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE;
    let defaultZipFile;

    beforeAll(async () => {
        expect(
            async () => await getZip(singleQueryDefaultTemplate)
        ).not.toThrow();

        defaultZipFile = await getZip(singleQueryDefaultTemplate);
    });

    test("Default template is a valid zip file", async () => {
        expect(defaultZipFile).toBeTruthy();
    });

    test("DataMashup XML exists, and valid PQ Base64 can be extracted", async () => {
        expect(
            async () => await pqUtils.getDataMashupFile(defaultZipFile)
        ).not.toThrowError();

        const { found, path, value } = await pqUtils.getDataMashupFile(
            defaultZipFile
        );

        expect(found).toBeTruthy();
        expect(value).toEqual(pqEmptySingleQueryBase64);
        expect(path).toEqual(item1Path);
    });

    test("ConnectedWorkbook XML exists as item2.xml", async () => {
        const { found, path, xmlString } = await pqUtils.getCustomXmlFile(
            defaultZipFile,
            URLS.CONNECTED_WORKBOOK,
            "UTF-8"
        );

        expect(found).toBeTruthy();
        expect(xmlString).toEqual(connectedWorkbookXmlMock);
        expect(path).toEqual(item2Path);
    });

    test("A single blank query named Query1 exists", async () => {
        const handler = new MashupHandler() as any;
        const base64Str = await pqUtils.getBase64(defaultZipFile);
        const { packageOPC } = handler.getPackageComponents(base64Str);
        const packageZip = await JSZip.loadAsync(packageOPC);
        const section1m: string = await handler.getSection1m(packageZip);
        const hasQuery1 = section1m.includes("Query1");

        expect(hasQuery1).toBeTruthy();
        expect(section1m).toEqual(section1mBlankQueryMock);
    });
});
