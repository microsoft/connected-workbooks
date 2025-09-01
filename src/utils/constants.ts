// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export const blobFileType = "blob";
export const uint8ArrayType = "uint8array";
export const application = "application/xlsx";
export const textResultType = "text";
export const xmlTextResultType = "text/xml";
export const pivotCachesPathPrefix = "pivotCacheDefinition";
export const trueValue = "1";
export const falseValue = "0";
export const emptyValue = "";
export const section1PathPrefix = "Section1/";
export const divider = "/";
export const maxQueryLength = 80;
export const trueStr = "true";
export const falseStr = "false";
export const BOM = "\ufeff";
export const maxCellCharacters = 32767;

export const connectionsXmlPath = "xl/connections.xml";
export const sharedStringsXmlPath = "xl/sharedStrings.xml";
export const sheetsXmlPath = "xl/worksheets/sheet1.xml";
export const tableXmlPath = "xl/tables/table1.xml";
export const queryTableXmlPath = "xl/queryTables/queryTable1.xml";
export const workbookXmlPath = "xl/workbook.xml";
export const queryTablesPath = "xl/queryTables/";
export const tablesFolderPath = "xl/tables/";
export const pivotCachesPath = "xl/pivotCache/";
export const section1mPath = "Formulas/Section1.m";
export const docPropsCoreXmlPath = "docProps/core.xml";
export const relsXmlPath = "_rels/.rels";
export const docMetadataXmlPath = "docMetadata";
export const docPropsRootElement = "cp:coreProperties";
export const workbookRelsXmlPath = "xl/_rels/workbook.xml.rels";
export const labelInfoXmlPath = "docMetadata/LabelInfo.xml";
export const docPropsAppXmlPath = "docProps/app.xml";

export const sharedStringsNotFoundErr = "SharedStrings were not found in template";
export const connectionsNotFoundErr = "Connections were not found in template";
export const WorkbookNotFoundERR = "workbook was not found in template";
export const sheetsNotFoundErr = "Sheets were not found in template";
export const base64NotFoundErr = "Base64 was not found in template";
export const emptyQueryMashupErr = "Query mashup is empty";
export const queryNameNotFoundErr = "Query name was not found";
export const queryAndPivotTableNotFoundErr = "No such query found in Query Table or Pivot Table found in given template";
export const queryConnectionNotFoundErr = "No connection found for query";
export const formulaSectionNotFoundErr = "Formula section wasn't found in template";
export const templateWithInitialDataErr = "Cannot use a template file with initial data";
export const queryTableNotFoundErr = "Query table wasn't found in template";
export const tableNotFoundErr = "Table wasn't found in template";
export const tableReferenceNotFoundErr = "Reference not found in the table XML.";
export const invalidValueInColumnErr = "Invalid cell value in column";
export const headerNotFoundErr = "Invalid JSON file, header is missing";
export const invalidDataTypeErr = "Invalid JSON file, invalid data type";
export const QueryNameMaxLengthErr = "Query names are limited to 80 characters";
export const QueryNameInvalidCharsErr = 'Query names cannot contain periods or quotation marks. (. ")';
export const EmptyQueryNameErr = "Query name cannot be empty";
export const stylesNotFoundErr = "Styles were not found in template";
export const InvalidColumnNameErr = "Invalid column name";
export const promotedHeadersCannotBeUsedWithoutAdjustingColumnNamesErr = "Headers cannot be promoted without adjusting column names";
export const unexpectedErr = "Unexpected error";
export const arrayIsntMxNErr = "Array isn't MxN";
export const relsNotFoundErr = ".rels were not found in template";
export const xlRelsNotFoundErr = "workbook.xml.rels were not found xl";
export const columnIndexOutOfRangeErr = "Column index out of range";
export const relationshipErr = "Relationship not found";
export const invalidCellValueErr = "Cell content exceeds maximum length of "  + maxCellCharacters+ " characters";

export const element = {
    sharedStringTable: "sst",
    text: "t",
    sharedStringItem: "si",
    cellValue: "v",
    databaseProperties: "dbPr",
    queryTable: "queryTable",
    cacheSource: "cacheSource",
    item: "Item",
    items: "Items",
    itemPath: "ItemPath",
    itemType: "ItemType",
    itemLocation: "ItemLocation",
    entry: "Entry",
    stableEntries: "StableEntries",
    tableColumns: "tableColumns",
    tableColumn: "tableColumn",
    table: "table",
    autoFilter: "autoFilter",
    definedName: "definedName",
    queryTableFields: "queryTableFields",
    queryTableField: "queryTableField",
    queryTableRefresh: "queryTableRefresh",
    sheetData: "sheetData",
    row: "row",
    dimension: "dimension",
    selection: "selection",
    kindCell: "c",
    sheet: "sheet",
    relationships: "Relationships",
    relationship: "Relationship"
};

export const elementAttributes = {
    connection: "connection",
    command: "command",
    refreshOnLoad: "refreshOnLoad",
    count: "count",
    uniqueCount: "uniqueCount",
    queryTable: "queryTable",
    connectionId: "connectionId",
    cacheSource: "cacheSource",
    name: "name",
    description: "description",
    id: "id",
    Id: "Id",
    relationId: "r:id",
    relationId1: "RId1",
    relationId2: "RId2",
    relationId3: "RId3",
    type: "Type",
    value: "Value",
    relationshipInfo: "RelationshipInfoContainer",
    resultType: "ResultType",
    fillColumnNames: "FillColumnNames",
    fillTarget: "FillTarget",
    fillLastUpdated: "FillLastUpdated",
    day: "d",
    uniqueName: "uniqueName",
    queryTableFieldId: "queryTableFieldId",
    reference: "ref",
    sqref: "sqref",
    tableColumnId: "tableColumnId",
    nextId: "nextId",
    row: "r",
    spans: "spans",
    x14acDyDescent: "x14ac:dyDescent",
    xr3uid: "xr3:uid",
    space: "xml:space",
    target: "Target",
};

export const dataTypeKind = {
    string: "str",
    number: "1",
    boolean: "b",
};

export const elementAttributesValues = {
    connectionName: (queryName: string) => `Query - ${queryName}`,
    connectionDescription: (queryName: string) => `Connection to the '${queryName}' query in the workbook.`,
    connection: (queryName: string) => `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="${queryName}";`,
    connectionCommand: (queryName: string) => `SELECT * FROM [${queryName}]`,
    tableResultType: () => "sTable",
};

export const defaults = {
    queryName: "Query1",
    sheetName: "Sheet1",
    columnName: "Column",
    tableName: "Table1",
};

export const URLS = {
    PQ: [
        "http://schemas.microsoft.com/DataMashup",
        "http://schemas.microsoft.com/DataExplorer",
        "http://schemas.microsoft.com/DataMashup/Temp",
        "http://schemas.microsoft.com/DataExplorer/Temp",
    ],
    CONNECTED_WORKBOOK: "http://schemas.microsoft.com/ConnectedWorkbook",
};

// Content-Type header to indicate that the content is an Excel document
export const headers = {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

export const OFU = {
    ViewUrl: "https://view.officeapps.live.com/op/view.aspx?src=http://connectedWorkbooks.excel/",
    PostUrl: "https://view.officeapps.live.com/op/viewpost.aspx?src=http://connectedWorkbooks.excel/",
    AllowTyping: "AllowTyping",
    WdOrigin: "wdOrigin",
    OpenInExcelOririgin: "OpenInExcel",
};
