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
export const contentTypesXmlPath = "[Content_Types].xml";
export const customXmlXmlPath = "customXml";

export const Errors = {
    sharedStringsNotFound: "SharedStrings were not found in template",
    connectionsNotFound: "Connections were not found in template",
    workbookNotFound: "workbook was not found in template",
    sheetsNotFound: "Sheets were not found in template",
    base64NotFound: "Base64 was not found in template",
    emptyQueryMashup: "Query mashup is empty",
    queryNameNotFound: "Query name was not found",
    queryAndPivotTableNotFound: "No such query found in Query Table or Pivot Table found in given template",
    queryConnectionNotFound: "No connection found for query",
    formulaSectionNotFound: "Formula section wasn't found in template",
    queryTableNotFound: "Query table wasn't found in template",
    tableNotFound: "Table wasn't found in template",
    tableReferenceNotFound: "Reference not found in the table XML.",
    invalidValueInColumn: "Invalid cell value in column",
    headerNotFound: "Invalid JSON file, header is missing",
    invalidDataType: "Invalid JSON file, invalid data type",
    queryNameMaxLength: "Query names are limited to 80 characters",
    queryNameInvalidChars: 'Query names cannot contain periods or quotation marks. (. ")',
    emptyQueryName: "Query name cannot be empty",
    stylesNotFound: "Styles were not found in template",
    invalidColumnName: "Invalid column name",
    promotedHeadersCannotBeUsedWithoutAdjustingColumnNames: "Headers cannot be promoted without adjusting column names",
    unexpected: "Unexpected error",
    arrayIsntMxN: "Array isn't MxN",
    relsNotFound: ".rels were not found in template",
    xlRelsNotFound: "workbook.xml.rels were not found xl",
    columnIndexOutOfRange: "Column index out of range",
    relationship: "Relationship not found",
    contentTypesNotFound: "contentTypes was not found in file",
    contentTypesParse: "Failed to parse [Content_Types].xml: Invalid XML structure",
    contentTypesElementNotFound: "contentTypes element was not found in parsed document",
    workbookRelsParse: "Failed to parse workbook relationships XML: Invalid XML structure",
    xmlParse: "Failed to parse XML: Parser error detected",
    relsParse: "Failed to parse .rels XML",
    connectionsParse: "Failed to parse connections XML",
    sharedStringsParse: "Failed to parse shared strings XML",
    worksheetParse: "Failed to parse worksheet XML",
    queryTableParse: "Failed to parse query table XML",
    pivotTableParse: "Failed to parse pivot table XML",
    workbookParse: "Failed to parse workbook XML",
    tableParse: "Failed to parse table XML",
    tablePathParse: "Failed to parse table XML for",
    invalidCellValueErr: "Cell content exceeds maximum length of "  + maxCellCharacters+ " characters",
};

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
    override: "Override",
    relationship: "Relationship",
    relationships: "Relationships",
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
    partName: "PartName",
    contentType: "ContentType",
    relationshipIdPrefix: "rId",
};

export const dataTypeKind = {
    string: "str",
    number: "n",
    boolean: "b",
};

export const elementAttributesValues = {
    connectionName: (queryName: string) => `Query - ${queryName}`,
    connectionDescription: (queryName: string) => `Connection to the '${queryName}' query in the workbook.`,
    connection: (queryName: string) => `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="${queryName}";`,
    connectionCommand: (queryName: string) => `SELECT * FROM [${queryName.replace(/]/g, ']]')}]`,
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

export const customXML = {
    customXMLItemContent: `<?xml version="1.0" encoding="utf-8"?><ConnectedWorkbook xmlns="http://schemas.microsoft.com/ConnectedWorkbook" version="1.0.0"></ConnectedWorkbook>`,
    customXMLItemPropsContent: `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<ds:datastoreItem ds:itemID="{0B384C3C-E1D4-401B-8CF4-6285949D7671}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.microsoft.com/ConnectedWorkbook"/></ds:schemaRefs></ds:datastoreItem>`,
    connectedWorkbookTag: '<ConnectedWorkbook',
    itemNumberPattern: /item(\d+)\.xml$/,
    itemFilePattern: /^item\d+\.xml$/,
    itemPropsPartNameTemplate: (itemIndex: string) => `/customXml/itemProps${itemIndex}.xml`,
    contentType: "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
    itemPathTemplate: (itemNumber: number | string) => `customXml/item${itemNumber}.xml`,
    itemPropsPathTemplate: (itemNumber: number | string) => `customXml/itemProps${itemNumber}.xml`,
    itemRelsPathTemplate: (itemNumber: number | string) => `customXml/_rels/item${itemNumber}.xml.rels`,
    customXMLRelationships: (itemNumber: number | string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps${itemNumber}.xml"/></Relationships>`,
    relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
    relativeItemPathTemplate: (itemNumber: number | string) => `../customXml/item${itemNumber}.xml`,
}