import workbookTemplate from "../src/workbookTemplate";
import { WorkbookManager }  from "../src/workbookManager";
import { sheetsXmlPath, workbookXmlPath, tableXmlPath, queryTableXmlPath } from "../src/constants";
import { dataTypes } from "../src/types";
import { sheetsXmlMock, workbookXmlMock, queryTableMock,  addZeroSheetsXmlMock } from "./mocks";
import JSZip from "jszip";

describe("Workbook Manager tests", () => {
    const workbookManager = new WorkbookManager() as any;
    test("test initial data in SheetsXML", async () => {
        const defaultZipFile = await JSZip.loadAsync(workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        await workbookManager.updateSheetsInitialData(defaultZipFile, {columnNames: ['Column1', 'Column2'], columnTypes: [dataTypes.string, dataTypes.number], 
                data: [['Column1', 'Column2'], ['1', '2']]});
        const sheetsXmlString = await defaultZipFile.file(sheetsXmlPath)?.async("text");
        if (sheetsXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        expect(sheetsXmlString).toContain(sheetsXmlMock);
        await workbookManager.updateSheetsInitialData(defaultZipFile, {columnNames: ['Column1', 'Column2'], columnTypes: [dataTypes.string, dataTypes.number],
                 data: [['Column1', 'Column2'], ['one', 'one'], ["two", "2"]]});
        const zeroSheetsXmlString = await defaultZipFile.file(sheetsXmlPath)?.async("text");
        expect(zeroSheetsXmlString).toContain(addZeroSheetsXmlMock);    
    })

    test("tests worksheetXML contains initial data dimensions", async () => {
        const defaultZipFile = await JSZip.loadAsync(workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        await workbookManager.updateWorkbookInitialData(defaultZipFile, {columnNames: ['Column1', 'Column2'], columnTypes: [dataTypes.string, dataTypes.number], 
                data: [['Column1', 'Column2'], ['1', '2']]});
        const worksheetXml = await defaultZipFile.file(workbookXmlPath)?.async("text");
        expect(worksheetXml).toContain(workbookXmlMock);
    })

    test("tests Pivot Tables contain initial data", async () => {
        const defaultZipFile = await JSZip.loadAsync(workbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true });
        await workbookManager.updatePivotTablesInitialData(defaultZipFile, {columnNames: ['Column1', 'Column2'], columnTypes: [dataTypes.string, dataTypes.number], 
                data: [['Column1', 'Column2'], ['1', '2']]});
        const tableXmlSheet = await defaultZipFile.file(tableXmlPath)?.async("text");
        expect(tableXmlSheet).toContain('count="2"');
        expect(tableXmlSheet).toContain('ref="A1:B2');
        expect(tableXmlSheet).toContain('uniqueName="1" name="Column1" queryTableFieldId="1"');
        expect(tableXmlSheet).toContain('uniqueName="2" name="Column2" queryTableFieldId="2"'); 
        const queryTableXmlSheet = await defaultZipFile.file(queryTableXmlPath)?.async("text");
        expect(queryTableXmlSheet).toContain(queryTableMock);
    })

});