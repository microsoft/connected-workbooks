import workbookTemplate from "../src/workbookTemplate";
import { WorkbookManager }  from "../src/workbookManager";
import { connectionsXmlPath } from "../src/constants";
import JSZip from "jszip";

test("Connection XML attributes contain new query name", async () => {
    const workbookManager = new WorkbookManager() as any;
    const defaultZipFile = await JSZip.loadAsync(workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true }) 
    await workbookManager.updateSingleQueryAttributes(defaultZipFile, "newQueryName", true);
    const connectionsXmlString = await defaultZipFile.file(connectionsXmlPath)?.async("text");
    const hasQueryNewName = connectionsXmlString?.includes("newQueryName");
    const hasQuery1 = connectionsXmlString?.includes("Query1");
    expect(hasQueryNewName).toBeTruthy();
    expect(hasQuery1).toBeFalsy;
    });