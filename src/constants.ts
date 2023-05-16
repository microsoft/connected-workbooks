// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const connectionsXmlPath = "xl/connections.xml";
export const sharedStringsXmlPath = "xl/sharedStrings.xml";
export const sheetsXmlPath = "xl/worksheets/sheet1.xml"
export const tableXmlPath = "xl/tables/table1.xml";
export const queryTableXmlPath = "xl/queryTables/queryTable1.xml";
export const workbookXmlPath = "xl/workbook.xml";
export const queryTablesPath = "xl/queryTables/";
export const pivotCachesPath = "xl/pivotCache/";
export const section1mPath = "Formulas/Section1.m";
export const docPropsCoreXmlPath = "docProps/core.xml";
export const docPropsRootElement = "cp:coreProperties";
export const stylesXmlPath = "xl/styles.xml";

export const sharedStringsNotFoundErr = "SharedStrings were not found in template";
export const connectionsNotFoundErr = "Connections were not found in template";
export const sheetsNotFoundErr = "Sheets were not found in template";
export const base64NotFoundErr = "Base64 was not found in template";
export const emptyQueryMashupErr = "Query mashup is empty";
export const queryAndPivotTableNotFoundErr = "No such query found in Query Table or Pivot Table found in given template";
export const queryConnectionNotFoundErr = "No connection found for query";
export const formulaSectionNotFoundErr = "Formula section wasn't found in template";
export const templateWithInitialDataErr = "Cannot receive template file with initial data";
export const queryTableNotFoundErr = "Query table wasn't found in template";
export const tableNotFoundErr = "Table wasn't found in template";
export const GridNotFoundErr = "Invalid JSON file, grid data is missing";
export const invalidValueInColumnErr = "Invalid cell value in column";
export const headerNotFoundErr = "Invalid JSON file, header is missing";
export const invalidDataTypeErr = "Invalid JSON file, invalid data type";
export const invalidFormatTypeErr = "Invalid JSON file, invalid format type";
export const stylesNotFoundErr = "Styles were not found in template";
export const invalidMissingFormatFromDateTimeErr = "Invalid JSON file, missing format from dateTime";

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
    differentialFormats: "dxfs",
    differentialFormat: "dxf",
    numberFormat: "numFmt",
    cellFormats: "cellXfs",
    cellFormat: "xf"
}

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
    tableColumnId: "tableColumnId",
    nextId: "nextId",
    row: "r",
    spans: "spans",
    x14acDyDescent: "x14ac:dyDescent",
    numberFormatId: "numFmtId",
    formatCode: "formatCode",
    dataDiffFormatId: "dataDxfId",
    fontId: "fontId",
    fillId: "fillId",
    borderId: "borderId",
    formatId: "xfId",
    applyNumberFormat: "applyNumberFormat"
};


export const elementAttributesValues = {
    connectionName: (queryName: string) => `Query - ${queryName}`,
    connectionDescription: (queryName: string) => `Connection to the '${queryName}' query in the workbook.`,
    connection: (queryName: string) => `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=${queryName};`,
    connectionCommand: (queryName: string) => `SELECT * FROM [${queryName}]`,
    tableResultType: () => "sTable"

}

export const milliSecPerDay = 86400000;
//This contains the number of days between 01/01/1970 and 01/01/1900
export const numberOfDaysTillExcelBeginYear = 25569;
export const monthsbeforeLeap = 2;
export const beginYear = 1900;
export const defaults = {
    queryName: "Query1",
};

export const dateFormatsRegex: { [key: string]: RegExp } = {
  "m/d/yyyy h:mm": /^([1-9]|0[1-9]|1[0-2])\/([1-9]|[012][0-9]|3[01])\/\d{4} ([01]\d|2[0-3]):([0-5]\d)$/
};

export const dateFormats: { [key: string]: number } = {
  "m/d/yyyy h:mm": 27
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
