// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { DocProps, DocPropsAutoUpdatedElements, DocPropsModifiableElements } from "../types";
import {
    docPropsCoreXmlPath,
    trueValue,
    falseValue,
    xmlTextResultType,
    element,
    elementAttributes,
    elementAttributesValues,
    connectionsNotFoundErr,
    sharedStringsNotFoundErr,
    queryTablesPath,
    textResultType,
    pivotCachesPath,
    pivotCachesPathPrefix,
    queryAndPivotTableNotFoundErr,
    emptyValue,
} from "./constants";
import documentUtils from "./documentUtils";
import { v4 } from "uuid";

const updateDocProps = async (zip: JSZip, docProps: DocProps = {}): Promise<void> => {
    const { doc, properties } = await documentUtils.getDocPropsProperties(zip);

    //set auto updated elements
    const docPropsAutoUpdatedElementsArr: ("created" | "modified")[] = Object.keys(DocPropsAutoUpdatedElements) as Array<
        keyof typeof DocPropsAutoUpdatedElements
    >;

    const nowTime: string = new Date().toISOString();

    docPropsAutoUpdatedElementsArr.forEach((tag) => {
        documentUtils.createOrUpdateProperty(doc, properties, DocPropsAutoUpdatedElements[tag], nowTime);
    });

    //set modifiable elements
    const docPropsModifiableElementsArr = Object.keys(DocPropsModifiableElements) as Array<keyof typeof DocPropsModifiableElements>;

    docPropsModifiableElementsArr
        .map((key) => ({
            name: DocPropsModifiableElements[key],
            value: docProps[key],
        }))
        .forEach((kvp) => {
            documentUtils.createOrUpdateProperty(doc, properties, kvp.name!, kvp.value);
        });

    const serializer: XMLSerializer = new XMLSerializer();
    const newDoc: string = serializer.serializeToString(doc);
    zip.file(docPropsCoreXmlPath, newDoc);
};

const updateConnections = (
    connectionsXmlString: string,
    queryName: string,
    refreshOnOpen: boolean
): { connectionId: string | undefined; connectionXmlFileString: string } => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const refreshOnLoadValue: string = refreshOnOpen ? trueValue : falseValue;
    const connectionsDoc: Document = parser.parseFromString(connectionsXmlString, xmlTextResultType);
    const connectionsProperties: HTMLCollectionOf<Element> = connectionsDoc.getElementsByTagName(element.databaseProperties);
    const dbPr: Element = connectionsProperties[0];
    dbPr.setAttribute(elementAttributes.refreshOnLoad, refreshOnLoadValue);

    // Update query details to match queryName
    dbPr.parentElement?.setAttribute(elementAttributes.name, elementAttributesValues.connectionName(queryName));
    dbPr.parentElement?.setAttribute(elementAttributes.description, elementAttributesValues.connectionDescription(queryName));
    dbPr.setAttribute(elementAttributes.connection, elementAttributesValues.connection(queryName));
    dbPr.setAttribute(elementAttributes.command, elementAttributesValues.connectionCommand(queryName));
    const connectionId: string | null | undefined = dbPr.parentElement?.getAttribute(elementAttributes.id);
    const connectionXmlFileString: string = serializer.serializeToString(connectionsDoc);

    if (connectionId === null) {
        throw new Error(connectionsNotFoundErr);
    }

    return { connectionId, connectionXmlFileString };
};

const updateSharedStrings = (sharedStringsXmlString: string, queryName: string): { sharedStringIndex: number; newSharedStrings: string } => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const sharedStringsDoc: Document = parser.parseFromString(sharedStringsXmlString, xmlTextResultType);
    const sharedStringsTable: Element = sharedStringsDoc.getElementsByTagName(element.sharedStringTable)[0];
    if (!sharedStringsTable) {
        throw new Error(sharedStringsNotFoundErr);
    }

    const textElementCollection: HTMLCollectionOf<Element> = sharedStringsDoc.getElementsByTagName(element.text);
    let textElement: Element | null = null;
    let sharedStringIndex: number = textElementCollection.length;
    if (textElementCollection && textElementCollection.length) {
        for (let i = 0; i < textElementCollection.length; i++) {
            if (textElementCollection[i].innerHTML === queryName) {
                textElement = textElementCollection[i];
                sharedStringIndex = i + 1;
                break;
            }
        }
    }

    if (textElement === null) {
        if (sharedStringsDoc.documentElement.namespaceURI) {
            textElement = sharedStringsDoc.createElementNS(sharedStringsDoc.documentElement.namespaceURI, element.text);
            textElement.textContent = queryName;
            const siElement: Element = sharedStringsDoc.createElementNS(sharedStringsDoc.documentElement.namespaceURI, element.sharedStringItem);
            siElement.appendChild(textElement);
            sharedStringsDoc.getElementsByTagName(element.sharedStringTable)[0].appendChild(siElement);
        }

        const value: string | null = sharedStringsTable.getAttribute(elementAttributes.count);
        if (value) {
            sharedStringsTable.setAttribute(elementAttributes.count, (parseInt(value) + 1).toString());
        }

        const uniqueValue: string | null = sharedStringsTable.getAttribute(elementAttributes.uniqueCount);
        if (uniqueValue) {
            sharedStringsTable.setAttribute(elementAttributes.uniqueCount, (parseInt(uniqueValue) + 1).toString());
        }
    }
    const newSharedStrings: string = serializer.serializeToString(sharedStringsDoc);

    return { sharedStringIndex, newSharedStrings };
};

const updateWorksheet = (sheetsXmlString: string, sharedStringIndex: string): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, xmlTextResultType);
    const cellValue: Element = sheetsDoc.getElementsByTagName(element.cellValue)[0];
    cellValue.innerHTML = sharedStringIndex;
    const newSheet: string = serializer.serializeToString(sheetsDoc);

    return newSheet;
};

const updatePivotTablesandQueryTables = async (zip: JSZip, queryName: string, refreshOnOpen: boolean, connectionId: string): Promise<void> => {
    // Find Query Table
    let found = false;
    const queryTablePromises: Promise<{
        path: string;
        queryTableXmlString: string;
    }>[] = [];
    zip.folder(queryTablesPath)?.forEach(async (relativePath, queryTableFile) => {
        queryTablePromises.push(
            (() => {
                return queryTableFile.async(textResultType).then((queryTableString) => {
                    return {
                        path: relativePath,
                        queryTableXmlString: queryTableString,
                    };
                });
            })()
        );
    });

    (await Promise.all(queryTablePromises)).forEach(({ path, queryTableXmlString }) => {
        const { isQueryTableUpdated, newQueryTable } = updateQueryTable(queryTableXmlString, connectionId, refreshOnOpen);
        zip.file(queryTablesPath + path, newQueryTable);
        if (isQueryTableUpdated) {
            found = true;
        }
    });
    if (found) {
        return;
    }

    // Find Pivot Table
    const pivotCachePromises: Promise<{
        path: string;
        pivotCacheXmlString: string;
    }>[] = [];

    zip.folder(pivotCachesPath)?.forEach(async (relativePath, pivotCacheFile) => {
        if (relativePath.startsWith(pivotCachesPathPrefix)) {
            pivotCachePromises.push(
                (() => {
                    return pivotCacheFile.async(textResultType).then((pivotCacheString) => {
                        return {
                            path: relativePath,
                            pivotCacheXmlString: pivotCacheString,
                        };
                    });
                })()
            );
        }
    });
    (await Promise.all(pivotCachePromises)).forEach(({ path, pivotCacheXmlString }) => {
        const { isPivotTableUpdated, newPivotTable } = updatePivotTable(pivotCacheXmlString, connectionId, refreshOnOpen);
        zip.file(pivotCachesPath + path, newPivotTable);
        if (isPivotTableUpdated) {
            found = true;
        }
    });
    if (!found) {
        throw new Error(queryAndPivotTableNotFoundErr);
    }
};

const updateQueryTable = (tableXmlString: string, connectionId: string, refreshOnOpen: boolean): { isQueryTableUpdated: boolean; newQueryTable: string } => {
    const refreshOnLoadValue: string = refreshOnOpen ? trueValue : falseValue;
    let isQueryTableUpdated = false;
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const queryTableDoc: Document = parser.parseFromString(tableXmlString, xmlTextResultType);
    const queryTable: Element = queryTableDoc.getElementsByTagName(element.queryTable)[0];
    let newQueryTable: string = emptyValue;
    if (queryTable.getAttribute(elementAttributes.connectionId) == connectionId) {
        queryTable.setAttribute(elementAttributes.refreshOnLoad, refreshOnLoadValue);
        newQueryTable = serializer.serializeToString(queryTableDoc);
        isQueryTableUpdated = true;
    }

    return { isQueryTableUpdated, newQueryTable };
};

const updatePivotTable = (tableXmlString: string, connectionId: string, refreshOnOpen: boolean): { isPivotTableUpdated: boolean; newPivotTable: string } => {
    const refreshOnLoadValue: string = refreshOnOpen ? trueValue : falseValue;
    let isPivotTableUpdated = false;
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const pivotCacheDoc: Document = parser.parseFromString(tableXmlString, xmlTextResultType);
    let cacheSource: Element = pivotCacheDoc.getElementsByTagName(element.cacheSource)[0];
    let newPivotTable: string = emptyValue;
    if (cacheSource.getAttribute(elementAttributes.connectionId) == connectionId) {
        cacheSource = cacheSource.parentElement!;
        cacheSource.setAttribute(elementAttributes.refreshOnLoad, refreshOnLoadValue);
        newPivotTable = serializer.serializeToString(pivotCacheDoc);
        isPivotTableUpdated = true;
    }

    return { isPivotTableUpdated, newPivotTable };
};

const randomizeConnectionsUUID = (connectionsXmlString: string): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const connectionsDoc: Document = parser.parseFromString(connectionsXmlString, xmlTextResultType);
    const connection: Element = connectionsDoc.getElementsByTagName(element.connection)[0];
    connection.setAttribute(elementAttributes.xr16uid, "{" + v4().toUpperCase() + "}");
    const newConnections: string = serializer.serializeToString(connectionsDoc);

    return newConnections;
};

const randomizeWorksheetUUID = (worksheetXmlString: string): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const worksheetDoc: Document = parser.parseFromString(worksheetXmlString, xmlTextResultType);
    const worksheet: Element = worksheetDoc.getElementsByTagName(element.worksheet)[0];
    worksheet.setAttribute(elementAttributes.xruid, "{" + v4().toUpperCase() + "}");
    const newSheet: string = serializer.serializeToString(worksheetDoc);

    return newSheet;
};

const randomizeTableUUID = (tableXmlString: string): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const tableDoc: Document = parser.parseFromString(tableXmlString, xmlTextResultType);
    const table: Element = tableDoc.getElementsByTagName(element.table)[0];
    const tableUUID: string = "{" + v4().toUpperCase() + "}";
    table.setAttribute(elementAttributes.xruid, tableUUID);
    const autoFilter: Element = tableDoc.getElementsByTagName(element.autoFilter)[0];
    autoFilter.setAttribute(elementAttributes.xruid, tableUUID);
    const tableColumn: Element = tableDoc.getElementsByTagName(element.tableColumn)[0];
    tableColumn.setAttribute(elementAttributes.xr3uid, "{" + v4().toUpperCase() + "}");
    const newTable: string = serializer.serializeToString(tableDoc);

    return newTable;
};

const randomizeQueryTableUUID = (queryTableXmlString: string): string => {
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const queryTableDoc: Document = parser.parseFromString(queryTableXmlString, xmlTextResultType);
    const queryTable: Element = queryTableDoc.getElementsByTagName(element.queryTable)[0];
    queryTable.setAttribute(elementAttributes.xr16uid, "{" + v4().toUpperCase() + "}");
    const newQueryTable: string = serializer.serializeToString(queryTableDoc);

    return newQueryTable;
};

export default {
    updateDocProps,
    updateConnections,
    updateSharedStrings,
    updateWorksheet,
    updatePivotTablesandQueryTables,
    updateQueryTable,
    updatePivotTable,
    randomizeConnectionsUUID,
    randomizeWorksheetUUID,
    randomizeTableUUID,
    randomizeQueryTableUUID
};
