import workbookTemplate from "../src/workbookTemplate";
import { pqUtils } from "../src/utils";
import { URLS } from "../src/constants";
import MashupHandler from "../src/mashupDocumentParser";
import { section1mBlankQueryMock } from "./mocks";
import JSZip from "jszip";

describe("single query template tests", () => {
    const singleQueryDefaultTemplate =
        workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE;
    let defaultZipFile;
    beforeAll(async () => {
        defaultZipFile = await JSZip.loadAsync(singleQueryDefaultTemplate, {
            base64: true,
        });
    });

    test("Default template is a valid zip file", async () => {
        expect(defaultZipFile).toBeTruthy();
    });

    test("DataMashup XML exists, and valid PQ Base64 can be extracted", async () => {
        const { found, path } = await pqUtils.getCustumXmlFile(
            defaultZipFile,
            URLS.DATA_MASHUP
        );

        expect(found).toBeTruthy();
        expect(path).toEqual("customXml/item1.xml");

        expect(
            async () => await pqUtils.getBase64(defaultZipFile)
        ).not.toThrowError();

        const base64Str = await pqUtils.getBase64(defaultZipFile);

        expect(base64Str).toBeTruthy();
    });

    test("ConnectedWorkbook XML exists as item2.xml", async () => {
        const { found, path } = await pqUtils.getCustumXmlFile(
            defaultZipFile,
            URLS.CONNECTED_WORKBOOK,
            "UTF-8"
        );

        expect(found).toBeTruthy();
        expect(path).toEqual("customXml/item2.xml");
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
