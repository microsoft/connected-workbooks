// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const connectionsXmlPath = "xl/connections.xml";
export const workbookXmlPath = "xl/workbook.xml";
export const sheetsXmlPath = "xl/worksheets/sheet1.xml"
export const tableXmlPath = "xl/tables/table1.xml";
export const queryTableXmlPath = "xl/queryTables/queryTable1.xml";
export const queryTablesPath = "xl/queryTables/";
export const pivotCachesPath = "xl/pivotCache/";
export const section1mPath = "Formulas/Section1.m";
export const docPropsCoreXmlPath = "docProps/core.xml";
export const docPropsRootElement = "cp:coreProperties";

export const defaults = {
    queryName: "Query1",
    relationshipInfo: `s{"columnCount":1,"keyColumnNames":[],"queryRelationships":[],"columnIdentities":["Section1/Query1/AutoRemovedColumns1.{Query1,0}"],"ColumnCount":1,"KeyColumnNames":[],"ColumnIdentities":["Section1/Query1/AutoRemovedColumns1.{Query1,0}"],"RelationshipInfo":[]}`
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
