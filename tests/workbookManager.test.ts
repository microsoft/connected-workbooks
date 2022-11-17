import workbookTemplate from "../src/workbookTemplate";
import { WorkbookManager }  from "../src/workbookManager";
import { connectionsXmlPath, sharedStringsXmlPath } from "../src/constants";
import { sharedStringsXmlMock } from "./mocks";

import JSZip from "jszip";

describe("Workbook Manager tests", () => {
    const workbookManager = new WorkbookManager() as any;

    test("Connection XML attributes contain new query name", async () => {
        const defaultZipFile = await JSZip.loadAsync(workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        await workbookManager.updateSingleQueryAttributes(defaultZipFile, "newQueryName", true);
        const connectionsXmlString = await defaultZipFile.file(connectionsXmlPath)?.async("text");
        const hasQueryNewName = connectionsXmlString?.includes("newQueryName");
        const hasQuery1 = connectionsXmlString?.includes("Query1");
        expect(hasQueryNewName).toBeTruthy();
        expect(hasQuery1).toBeFalsy;
    })

    test("SharedStrings XML contains new query name", async () => {
        const defaultZipFile = await JSZip.loadAsync(workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true }) 
        const sharedStringId = await workbookManager.editSharedStrings(defaultZipFile, "newQueryName");
        expect(sharedStringId).toEqual(2);
        if (sharedStringsXmlMock) {
            const sharedStringsXmlString = await defaultZipFile.file(sharedStringsXmlPath)?.async("text");
            const mockSharedString = sharedStringsXmlMock.replace(/ /g, "");
            expect(sharedStringsXmlString?.replace(/ /g, "")).toContain(mockSharedString);
        }
    })
});