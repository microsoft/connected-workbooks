import { tableUtils } from "../src/utils";
import { queryTableMock, sheetsXmlMock, workbookXmlMock } from "./mocks";

describe("Table Utils tests", () => {
    test("tests workbookXML contains initial data dimensions", () => {
        const defaultString =
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="24729"/><workbookPr codeName="ThisWorkbook" defaultThemeVersion="166925"/><mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="C:Usersv-ahmadsbeihDesktop" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{93EF201C-7856-4B60-94D4-65DDB8F3F16A}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/><bookViews><workbookView xWindow="28680" yWindow="-120" windowWidth="29040" windowHeight="15990" xr2:uid="{DB915CB9-8DD9-492A-A471-C61E61200113}"/></bookViews><sheets><sheet name="Query1" sheetId="2" r:id="rId1"/><sheet name="Sheet1" sheetId="1" r:id="rId2"/></sheets><definedNames><definedName name="ExternalData_1" localSheetId="0" hidden="1">Sheet1!$A$1:$B$2</definedName></definedNames><calcPr calcId="191029"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"><x15:workbookPr chartTrackingRefBase="1"/></ext><ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"><xcalcf:calcFeatures><xcalcf:feature name="microsoft.com:RD"/><xcalcf:feature name="microsoft.com:FV"/><xcalcf:feature name="microsoft.com:LET_WF"/><xcalcf:feature name="microsoft.com:LAMBDA_WF"/></xcalcf:calcFeatures></ext></extLst></workbook>';
        const worksheetXml = tableUtils.updateWorkbookInitialData(defaultString, {
            columnNames: ["Column1", "Column2"],
            rows: [["1", "2"]],
        });
        expect(worksheetXml).toContain(workbookXmlMock);
    });

    test("tests Pivot Tables contain initial data", () => {
        const defaultString =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" id="1" xr:uid="{D8539CF6-04E5-464D-9950-5A36C5A1FCFE}" name="Query1" displayName="Query1" ref="A1:A2" tableType="queryTable" totalsRowShown="0"><autoFilter ref="A1:A2" xr:uid="{D8539CF6-04E5-464D-9950-5A36C5A1FCFE}"/><tableColumns count="1"><tableColumn id="1" xr3:uid="{D1084858-8AE5-4728-A9BE-FE78821CDFFF}" uniqueName="1" name="Query1" queryTableFieldId="1" dataDxfId="0"/></tableColumns><tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
        const tableXmlSheet = tableUtils.updateTablesInitialData(
            defaultString,
            {
                columnNames: ["Column1", "Column2"],
                rows: [["1", "2"]],
            },
            true
        );
        expect(tableXmlSheet).toContain('count="2"');
        expect(tableXmlSheet).toContain('ref="A1:B2');

        expect(tableXmlSheet).toContain('uniqueName="1"');
        expect(tableXmlSheet).toContain('name="Column1"');
        expect(tableXmlSheet).toContain('queryTableFieldId="1"');

        expect(tableXmlSheet).toContain('uniqueName="2"');
        expect(tableXmlSheet).toContain('name="Column2"');
        expect(tableXmlSheet).toContain('queryTableFieldId="2"');
    });

    test("tests blank Table contain initial data", () => {
        const defaultString =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" id="1" xr:uid="{D8539CF6-04E5-464D-9950-5A36C5A1FCFE}" name="Query1" displayName="Query1" ref="A1:A2" tableType="queryTable" totalsRowShown="0"><autoFilter ref="A1:A2" xr:uid="{D8539CF6-04E5-464D-9950-5A36C5A1FCFE}"/><tableColumns count="1"><tableColumn id="1" xr3:uid="{D1084858-8AE5-4728-A9BE-FE78821CDFFF}" uniqueName="1" name="Query1" queryTableFieldId="1" dataDxfId="0"/></tableColumns><tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
        const tableXmlSheet = tableUtils.updateTablesInitialData(
            defaultString,
            {
                columnNames: ["Column1", "Column2"],
                rows: [["1", "2"]],
            },
            false
        );
        expect(tableXmlSheet).toContain('count="2"');
        expect(tableXmlSheet).toContain('ref="A1:B2');
        expect(tableXmlSheet).toContain('name="Column2"');
        expect(tableXmlSheet).toContain('name="Column1"');

        // Not contains query table metadata.
        expect(tableXmlSheet).not.toContain('uniqueName="1"');
        expect(tableXmlSheet).not.toContain('queryTableFieldId="1"');

        expect(tableXmlSheet).not.toContain('uniqueName="2"');
        expect(tableXmlSheet).not.toContain('queryTableFieldId="2"');
    });

    test("test valid initial data in SheetsXML", () => {
        const defaultString =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{EDF0138E-D216-4CD1-8EFA-1396A1BB4478}"><sheetPr codeName="Sheet1"/><dimension ref="A1:A2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="14.4" x14ac:dyDescent="0.3"/><cols><col min="1" max="1" width="9.6640625" bestFit="1" customWidth="1"/></cols><sheetData><row r="1" spans="1:1" x14ac:dyDescent="0.3"><c r="A1" t="s"><v>0</v></c></row><row r="2" spans="1:1" x14ac:dyDescent="0.3"><c r="A2" t="s"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>';
        const sheetsXmlString = tableUtils.updateSheetsInitialData(defaultString, {
            columnNames: ["Column1", "Column2"],
            rows: [["1", "2"]],
        });
        expect(sheetsXmlString).toContain(sheetsXmlMock);
    });

    test("tests Query Tables contain initial data", () => {
        const defaultString =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16" xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16" name="ExternalData_1" connectionId="1" xr16:uid="{24C17B89-3CD3-4AA5-B84F-9FF5F35245D7}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0"><queryTableRefresh nextId="2"><queryTableFields count="1"><queryTableField id="1" name="Query1" tableColumnId="1"/></queryTableFields></queryTableRefresh></queryTable>';
        const queryTableXmlSheet = tableUtils.updateQueryTablesInitialData(defaultString, {
            columnNames: ["Column1", "Column2"],
            rows: [["1", "2"]],
        });
        expect(queryTableXmlSheet).toContain(queryTableMock);
    });

    test("tests Get next column Name", () => {
        const columnNames = ["Column", "Column (1)", "Column (2)", "Column (3)", "Banana"];
        expect(tableUtils.getNextAvailableColumnName(columnNames, "Column")).toEqual("Column (4)");
        expect(tableUtils.getNextAvailableColumnName(columnNames, "Banana")).toEqual("Banana (1)");
        expect(tableUtils.getNextAvailableColumnName(columnNames, "unexists")).toEqual("unexists");
    });

    test("tests Get adjusted column Name", () => {
        expect(tableUtils.getAdjustedColumnNames(["Column", 2, false, 2, "Banana"])).toEqual(["Column", "2", "false", "2 (1)", "Banana"]);
        expect(tableUtils.getAdjustedColumnNames(["Column", 2, "", 2, "Banana"])).toEqual(["Column", "2", "Column (1)", "2 (1)", "Banana"]);
        expect(tableUtils.getAdjustedColumnNames(["Column", "Column", "Column (2)", "Column (3)"])).toEqual([
            "Column",
            "Column (1)",
            "Column (2)",
            "Column (3)",
        ]);
    });

    test("tests Get raw column Name success", () => {
        expect(tableUtils.getRawColumnNames(["Column", "Column1", "Column2", 2])).toEqual(["Column", "Column1", "Column2", "2"]);
    });

    test("tests Get raw column Name empty name", () => {
        try {
            tableUtils.getRawColumnNames(["Column", 2, false, 2, "Banana"]);
        } catch (error) {
            expect(error).toBeTruthy();
        }
    });

    test("tests Get raw column Name conflict name", () => {
        try {
            tableUtils.getRawColumnNames(["Column", 2, "", 2, "Banana"]);
        } catch (error) {
            expect(error).toBeTruthy();
        }
    });
});
