export const connectedWorkbookXmlMock =
    '<?xml version="1.0" encoding="utf-8"?><ConnectedWorkbook xmlns="http://schemas.microsoft.com/ConnectedWorkbook" version="1.0.0"></ConnectedWorkbook>';
export const sheetsXmlMock =
    '<sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.3"><c r="A1" t="str"><v>Column1</v></c><c r="B1" t="str"><v>Column2</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.3"><c r=\"A2\" t=\"str\"><v>1</v></c><c r=\"B2\" t=\"1\"><v>2</v></c></row></sheetData>';
export const addZeroSheetsXmlMock =
    '<sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.3"><c r="A1" t="str"><v>Column1</v></c><c r="B1" t="str"><v>Column2</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.3"><c r="A2" t="str"><v>one</v></c><c r="B2" t="1"><v>0</v></c></row><row r="3" spans="1:2" x14ac:dyDescent="0.3"><c r="A3" t="str"><v>two</v></c><c r="B3" t="1"><v>2</v></c></row></sheetData>';
export const workbookXmlMock = 
    '<definedName name="ExternalData_1" localSheetId="0" hidden="1">Query1!$A$1:$B$2</definedName>';
export const queryTableMock = 
    '<queryTableRefresh nextId="3"><queryTableFields count="2"><queryTableField id="1" name="Column1" tableColumnId="1"/><queryTableField id="2" name="Column2" tableColumnId="2"/></queryTableFields></queryTableRefresh></queryTable>';