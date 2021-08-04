import MashupHandler from "../src/mashupDocumentParser";
import { zipUtils, arrayUtils, pqUtils } from "../src/utils";
import { simpleQuery, section1mMock } from "./mocks";
import base64 from "base64-js";

describe("Mashup Document Parser tests", () => {
    test("ReplaceSingleQuery test", async () => {
        const mashupHandler = new MashupHandler();

        const defaultZipFile = await zipUtils.loadAsyncDefaultTemplate();
        const base64str = await zipUtils.getBase64(defaultZipFile);

        const replacedQueryStr = await mashupHandler.ReplaceSingleQuery(
            base64str,
            simpleQuery
        );

        const buffer = base64.toByteArray(replacedQueryStr).buffer;
        const mashupArray = new arrayUtils.ArrayReader(buffer);
        const startArray = mashupArray.getBytes(4);
        const packageSize = mashupArray.getInt32();
        const packageOPC = mashupArray.getBytes(packageSize);

        const zip = await zipUtils.loadAsync(packageOPC);
        const section1m = await pqUtils.getSection1m(zip);

        expect(section1m.replace(/ /g, "")).toEqual(
            section1mMock.replace(/ /g, "")
        );
    });
});