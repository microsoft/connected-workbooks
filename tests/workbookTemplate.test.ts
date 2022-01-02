import workbookTemplate from "../src/workbookTemplate";
import { pqUtils } from "../src/utils";
import { pqCustomXmlPath } from "../src/constants";
import MashupHandler from "../src/mashupDocumentParser";
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

    test("Default template is valid", async () => {
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile);

        expect(defaultZipFile).toBeTruthy();
        expect(originalBase64Str).toBeTruthy();
    });

    test("ConnectedWorkbook XML exists", async () => {
        const connectedWorkbookXml = defaultZipFile.file(pqCustomXmlPath);

        expect(connectedWorkbookXml).toBeTruthy();
    });

    test("Base64 string exists and valid for parsing", async () => {
        expect(
            async () => await pqUtils.getBase64(defaultZipFile)
        ).not.toThrowError();

        const base64Str = await pqUtils.getBase64(defaultZipFile);

        expect(base64Str).toBeTruthy();
    });

    test("A single query exists", async () => {
        const handler = new MashupHandler() as any;
        const base64Str = await pqUtils.getBase64(defaultZipFile);
        const { packageOPC } = handler.getPackageComponents(base64Str);
        const packageZip = await JSZip.loadAsync(packageOPC);
        const section1m: string = await handler.getSection1m(packageZip);

        const hasQuery1 = section1m.includes("shared Query1");

        expect(hasQuery1).toBeTruthy();
    });
});
