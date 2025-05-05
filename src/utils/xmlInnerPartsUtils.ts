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
    docMetadataXmlPath,
    relsXmlPath,
    unexpectedErr,
    relsNotFoundErr,
    WorkbookNotFoundERR,
    workbookXmlPath,
    tableNotFoundErr,
    tableReferenceNotFoundErr,
} from "./constants";
import documentUtils from "./documentUtils";
import { DOMParser, XMLSerializer } from "xmldom-qsa";

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

const clearLabelInfo = async (zip: JSZip): Promise<void> => {
    // remove docMetadata folder that contains only LabelInfo.xml in template file.
    zip.remove(docMetadataXmlPath);

    // fix rels
    const relsString = await zip.file(relsXmlPath)?.async(textResultType);
    if (relsString === undefined) {
        throw new Error(relsNotFoundErr);
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(relsString, xmlTextResultType);
    const relationships = doc.querySelector("Relationships");
    if (relationships === null) {
        throw new Error(unexpectedErr);
    }
    const element = relationships.querySelector('Relationship[Target="docMetadata/LabelInfo.xml"]');
    if (element) {
        relationships.removeChild(element);
    }
    relationships.querySelector('Relationship[Target="xl/workbook.xml"]')?.setAttribute("Id", "rId1");
    relationships.querySelector('Relationship[Target="docProps/core.xml"]')?.setAttribute("Id", "rId2");
    relationships.querySelector('Relationship[Target="docProps/app.xml"]')?.setAttribute("Id", "rId3");

    const serializer: XMLSerializer = new XMLSerializer();
    const newDoc: string = serializer.serializeToString(doc);
    zip.file(relsXmlPath, newDoc);
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
    (dbPr.parentNode as Element)?.setAttribute(elementAttributes.name, elementAttributesValues.connectionName(queryName));
    (dbPr.parentNode as Element)?.setAttribute(elementAttributes.description, elementAttributesValues.connectionDescription(queryName));
    dbPr.setAttribute(elementAttributes.connection, elementAttributesValues.connection(queryName));
    dbPr.setAttribute(elementAttributes.command, elementAttributesValues.connectionCommand(queryName));
    const connectionId: string | null | undefined = (dbPr.parentNode as Element)?.getAttribute(elementAttributes.id);
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
            if (textElementCollection[i].textContent === queryName) {
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
    sheetsDoc.getElementsByTagName(element.cellValue)[0].innerHTML = sharedStringIndex.toString();
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

// get sheet name from workbook
const getSheetIdByNameFromZip = async (zip: JSZip, sheetName: string): Promise<string> => {
    const workbookXmlString: string | undefined = await zip.file(workbookXmlPath)?.async("text");
    if (!workbookXmlString) {
        throw new Error(WorkbookNotFoundERR);
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(workbookXmlString, xmlTextResultType);

    const sheetElements = doc.getElementsByTagName("sheet");
    for (let i = 0; i < sheetElements.length; i++) {
        if (sheetElements[i].getAttribute("name") === sheetName) {
            return (i + 1).toString();
        }
    }
    throw new Error(`Sheet with name ${sheetName} not found`);
};

// get definedName
const getReferenceFromTable = async (zip: JSZip, tablePath: string): Promise<string> => {
    const tableXmlString: string | undefined = await zip.file(tablePath)?.async("text");
    if (!tableXmlString) {
        throw new Error(WorkbookNotFoundERR);
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(tableXmlString, xmlTextResultType);
    const tableElements = doc.getElementsByTagName("table");
    const reference = tableElements[0]?.getAttribute("ref");
    if (!reference) {
        throw new Error(tableReferenceNotFoundErr);
    }

    return reference;
};

const findTablePathFromZip = async (zip: JSZip, targetTableName: string): Promise<string> => {
    const tablesFolder = zip.folder("xl/tables");
    if (!tablesFolder) return "";

    const tableFilePromises: Promise<{ path: string; content: string }>[] = [];
    tablesFolder.forEach((relativePath, file) => {
        tableFilePromises.push(
            file.async(textResultType).then(content => ({ path: relativePath, content }))
        );
    });

    const tableFiles = await Promise.all(tableFilePromises);
    const parser = new DOMParser();
    for (const { path, content } of tableFiles) {
        const doc = parser.parseFromString(content, xmlTextResultType);
        const tableElem = doc.getElementsByTagName("table")[0];
        if (tableElem && tableElem.getAttribute("name") === targetTableName) {
            return path;
        }
    }

    throw new Error(tableNotFoundErr);
};

export default {
    updateDocProps,
    clearLabelInfo,
    updateConnections,
    updateSharedStrings,
    updateWorksheet,
    updatePivotTablesandQueryTables,
    updateQueryTable,
    updatePivotTable,
    getSheetIdByNameFromZip,
    getReferenceFromTable,
    findTablePathFromZip,
};
