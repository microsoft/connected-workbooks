import MashupHandler from "../src/mashupDocumentParser";
import { arrayUtils, pqUtils } from "../src/utils";
import { simpleQueryMock, section1mSimpleQueryMock } from "./mocks";
import base64 from "base64-js";
import JSZip from "jszip";
import WorkbookTemplate from "../src/workbookTemplate";
import { section1mPath } from "../src/constants";
describe("Mashup Document Parser tests", () => {
    test("ReplaceSingleQuery test", async () => {
        const mashupHandler = new MashupHandler();

        const defaultZipFile = await JSZip.loadAsync(
            WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE,
            { base64: true }
        );
        const originalBase64Str = await pqUtils.getBase64(defaultZipFile);

        const replacedQueryBase64Str = await mashupHandler.ReplaceSingleQuery(
            originalBase64Str,
            simpleQueryMock
        );

        const buffer = base64.toByteArray(replacedQueryBase64Str).buffer;
        const mashupArray = new arrayUtils.ArrayReader(buffer);
        const startArray = mashupArray.getBytes(4);
        const packageSize = mashupArray.getInt32();
        const packageOPC = mashupArray.getBytes(packageSize);

        const zip = await JSZip.loadAsync(packageOPC);
        const section1m = await zip.file(section1mPath)?.async("text");

        expect(section1m.replace(/ /g, "")).toEqual(
            section1mSimpleQueryMock.replace(/ /g, "")
        );
    });
});
