// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const connectionsXmlPath = "xl/connections.xml";
export const sharedStringsXmlPath = "xl/sharedStrings.xml";
export const sheetsXmlPath = "xl/worksheets/sheet1.xml"
export const queryTablesPath = "xl/queryTables/";
export const pivotCachesPath = "xl/pivotCache/";
export const section1mPath = "Formulas/Section1.m";
export const docPropsCoreXmlPath = "docProps/core.xml";
export const docPropsRootElement = "cp:coreProperties";

export const defaults = {
    queryName: "Query1",
    connectionOnlyQueryName: "Query2"
};

export const elementAttributes = {
    connection: "connection",
    command: "command",
    refreshOnLoad: "refreshOnLoad",
    sharedStringTable: "sst",
    text: "t",
    sharedStringItem: "si",
    count: "count",
    uniqueCount: "uniqueCount",
    queryTable: "queryTable",
    connectionId: "connectionId",
    cacheSource: "cacheSource",
    name: "name",
    description: "description",
    id: "id",
    v: "v"
};


export const SHARED_STRINGS_NOT_FOUND = "SharedStrings were not found in template";
export const CONNECTIONS_NOT_FOUND = "Connections were not found in template";
export const SHEETS_NOT_FOUND = "Sheets were not found in template";
export const BASE64_NOT_FOUND = "Base64 was not found in template";
export const EMPTY_QUERY_MASHUP = "Query mashup is empty";
export const QUERY_TABLE_NOT_FOUND = "No Query Table or Pivot Table found for query in given template.";
export const QUERY_CONNECTION_NOT_FOUND = `No connection found for query`;

export const URLS = {
    PQ: [
        "http://schemas.microsoft.com/DataMashup",
        "http://schemas.microsoft.com/DataExplorer",
        "http://schemas.microsoft.com/DataMashup/Temp",
        "http://schemas.microsoft.com/DataExplorer/Temp",
    ],
    CONNECTED_WORKBOOK: "http://schemas.microsoft.com/ConnectedWorkbook",
};
