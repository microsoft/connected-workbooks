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
    relsNotFoundErr,
    WorkbookNotFoundERR,
    workbookXmlPath,
    tableNotFoundErr,
    tableReferenceNotFoundErr,
    workbookRelsXmlPath,
    xlRelsNotFoundErr,
    customXML,
    customXmlXmlPath,
    contentTypesNotFoundERR,
    contentTypesXmlPath,
    tablesFolderPath,
    labelInfoXmlPath,
    docPropsAppXmlPath,
    relationshipErr,
    contentTypesParseErr,
    contentTypesElementNotFoundERR,
    workbookRelsParseErr,
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

const removeLabelInfoRelationship = (doc: Document, relationships: Element) => {
    // Find and remove LabelInfo.xml relationship
    const relationshipElements = doc.getElementsByTagName(element.relationship);
    for (let i = 0; i < relationshipElements.length; i++) {
        const rel = relationshipElements[i];
        if (rel.getAttribute(elementAttributes.target) === labelInfoXmlPath) {
            relationships.removeChild(rel);
            break;
        }
    }
};

const updateRelationshipIds = (doc: Document) => {
    // Update relationship IDs
    const relationshipElements = doc.getElementsByTagName(element.relationship);
    for (let i = 0; i < relationshipElements.length; i++) {
        const rel = relationshipElements[i];
        const target = rel.getAttribute(elementAttributes.target);
        if (target === workbookXmlPath) {
            rel.setAttribute(elementAttributes.Id, elementAttributes.relationId1);
        } else if (target === docPropsCoreXmlPath) {
            rel.setAttribute(elementAttributes.Id, elementAttributes.relationId2);
        } else if (target === docPropsAppXmlPath) {
            rel.setAttribute(elementAttributes.Id, elementAttributes.relationId3);
        }
    }
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
    const relationshipsList = doc.getElementsByTagName(element.relationships);
    if (!relationshipsList || relationshipsList.length === 0) {
        throw new Error(relationshipErr);
    }

    const relationships = relationshipsList[0];
    if (!relationships) {
        throw new Error(relationshipErr);
    }

    removeLabelInfoRelationship(doc, relationships);
    updateRelationshipIds(doc);

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

/**
 * Retrieves the target path of a sheet from workbook relationships by its relationship ID.
 */
async function getSheetPathFromXlRelId(zip: JSZip, rId: string): Promise<string> {
    const relsFile = zip.file(workbookRelsXmlPath);
    if (!relsFile) {
        throw new Error(xlRelsNotFoundErr);
    }

    const relsString = await relsFile.async(textResultType);
    const relsDoc = new DOMParser().parseFromString(relsString, xmlTextResultType);

    // Avoid querySelector due to xmldom-qsa edge cases; iterate elements safely
    const relationships = relsDoc.getElementsByTagName("Relationship");
    let target: string | null = null;
    for (let i = 0; i < relationships.length; i++) {
        const el = relationships[i];
        if (el && el.getAttribute && el.getAttribute("Id") === rId) {
            target = el.getAttribute(elementAttributes.target);
            break;
        }
    }

    if (!target) {
        throw new Error(`Relationship not found or missing Target for Id: ${rId}`);
    }

    return target;
}

// get sheet name from workbook
const getSheetPathByNameFromZip = async (zip: JSZip, sheetName: string): Promise<string> => {
    const workbookXmlString: string | undefined = await zip.file(workbookXmlPath)?.async("text");
    if (!workbookXmlString) {
        throw new Error(WorkbookNotFoundERR);
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(workbookXmlString, xmlTextResultType);

    const sheetElements = doc.getElementsByTagName(element.sheet);
    for (let i = 0; i < sheetElements.length; i++) {
        if (sheetElements[i].getAttribute(elementAttributes.name) === sheetName) {
            const rId = sheetElements[i].getAttribute(elementAttributes.relationId);
            if (rId) {
                return getSheetPathFromXlRelId(zip, rId);
            }
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
    const tableElements = doc.getElementsByTagName(element.table);
    const reference = tableElements[0]?.getAttribute(elementAttributes.reference);
    if (!reference) {
        throw new Error(tableReferenceNotFoundErr);
    }

    return reference.split(":")[0]; // Return the start cell reference (e.g., "A1" from "A1:B10")
};

const findTablePathFromZip = async (zip: JSZip, targetTableName: string): Promise<string> => {
    const tablesFolder = zip.folder(tablesFolderPath); 
    if (!tablesFolder) return emptyValue;

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
        const tableElem = doc.getElementsByTagName(element.table)[0];
        if (tableElem && tableElem.getAttribute(elementAttributes.name) === targetTableName) {
            return path;
        }
    }

    throw new Error(tableNotFoundErr);
};

/**
 * Determines the next available item number for a custom XML item in the Excel workbook.
 * Scans the customXml folder to find existing item files and returns the next sequential number.
 * 
 * @param zip - The JSZip instance containing the Excel workbook structure
 * @returns Promise resolving to the next available item number (starting from 1 if no items exist)
 * 
 * @example
 * // If customXml folder contains item1.xml, item2.xml, returns 3
 * const nextNumber = await getCustomXmlItemNumber(zip);
 */
const getCustomXmlItemNumber = async (zip: JSZip): Promise<number> => {
    const customXmlFolder = zip.folder(customXmlXmlPath);
    if (!customXmlFolder) {
        return 1; // start from 1 if folder doesn't exist
    }

    // Build regex to match custom XML item files in the customXml folder
    const re = new RegExp(`^${customXmlXmlPath}/${customXML.itemNumberPattern.source}$`);
    const matches = zip.file(re);

    let max = 0;
    // Iterate through all matching files to find the highest item number
    for (const f of matches) {
        const m = f.name.match(customXML.itemNumberPattern);
        if (m) {
            const n = parseInt(m[1], 10); // Extract the number from the filename
            if (!Number.isNaN(n) && n > max) {
                max = n;
            }
        }
    }

    return max + 1; // Return next available number
};

/**
 * Checks if a custom XML item with connected-workbooks already exists in the Excel workbook.
 * Searches through all custom XML files in the customXml folder to find a match with the expected content.
 * 
 * @param zip - The JSZip instance containing the Excel workbook structure
 * @returns Promise resolving to true if the custom XML item exists, false otherwise
 * 
 * @example
 * const exists = await isCustomXmlExists(zip);
 * if (!exists) {
 *   // Add new custom XML item
 * }
 */
const isCustomXmlExists = async (zip: JSZip): Promise<boolean> => {
    const customXmlFolder = zip.folder(customXmlXmlPath);
    if (!customXmlFolder) {
        return false; // customXml folder does not exist
    }

    // Get all files matching the custom XML item pattern
    const customXmlFiles = customXmlFolder.file(customXML.itemFilePattern);
    for (const file of customXmlFiles) {
        try {
            const content = await file.async(textResultType);
            if (content.includes(customXML.connectedWorkbookTag)) {
                return true; // Found matching custom XML item
            }
        } catch (error) {
            // Skip files that can't be read and continue with the next file
            continue;
        }
    }

    return false; // No matching custom XML item found
};

/**
 * Adds a content type override entry to the [Content_Types].xml file for a custom XML item.
 * This registration is required for Excel to recognize and process the custom XML item.
 * 
 * @param zip - The JSZip instance containing the Excel workbook structure
 * @param itemIndex - The index/number of the custom XML item to register
 * @throws {Error} When the [Content_Types].xml file is not found or cannot be parsed
 * 
 * @example
 * await addToContentType(zip, "1"); // Registers customXml/item1.xml in content types
 */
const addToContentType = async (zip: JSZip, itemIndex: string) => {
    const contentTypesXmlString: string | undefined = await zip.file(contentTypesXmlPath)?.async(textResultType);
    if (!contentTypesXmlString) {
        throw new Error(contentTypesNotFoundERR);
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(contentTypesXmlString, xmlTextResultType);

    // Check if parsing was successful by verifying we have a valid document
    if (!doc || !doc.documentElement) {
        throw new Error(contentTypesParseErr);
    }

    const partName = customXML.itemPropsPartNameTemplate(itemIndex);
    const contentTypeValue = customXML.contentType;
    const typesElement = doc.documentElement;
    if (!typesElement) {
        throw new Error(contentTypesElementNotFoundERR);
    }

    const ns = doc.documentElement.namespaceURI;
    const overrideEl = ns ? doc.createElementNS(ns, element.override) : doc.createElement(element.override);
    overrideEl.setAttribute(elementAttributes.partName, partName);
    overrideEl.setAttribute(elementAttributes.contentType, contentTypeValue);
    typesElement.appendChild(overrideEl);
    
    const serializer = new XMLSerializer();
    const newDoc = serializer.serializeToString(doc);
    zip.file(contentTypesXmlPath, newDoc);
};

/**
 * Adds a relationship entry to the workbook relationships file for a custom XML item.
 * Creates a new relationship that links the workbook to the custom XML item.
 * 
 * @param zip - The JSZip instance containing the Excel workbook structure
 * @param itemIndex - The index/number of the custom XML item to create a relationship for
 * @throws {Error} When the workbook relationships file is not found or cannot be parsed
 * 
 * @example
 * await addCustomXmlToRels(zip, "1"); // Creates relationship to customXml/item1.xml
 */
const addCustomXmlToRels = async (zip: JSZip, itemIndex: string) => {
    const relsXmlString: string | undefined = await zip.file(workbookRelsXmlPath)?.async(textResultType);
    if (!relsXmlString) {
        throw new Error(relsNotFoundErr);
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXmlString, xmlTextResultType);
    
    // Check if parsing was successful by verifying we have a valid document
    if (!doc || !doc.documentElement) {
        throw new Error(workbookRelsParseErr);
    }
    
    // Use getElementsByTagName for better cross-platform compatibility
    const relationshipsElements = doc.getElementsByTagName(element.relationships);
    if (!relationshipsElements || relationshipsElements.length === 0) {
        throw new Error(relationshipErr);
    }

    const relationshipsElement = relationshipsElements[0];
    const highestRid = getHighestRelationshipId(relationshipsElement);

    // Generate new relationship details
    const newRid = `${elementAttributes.relationshipIdPrefix}${highestRid + 1}`;
    const target = customXML.relativeItemPathTemplate(itemIndex);
    const type = customXML.relationshipType;

    // Create and configure the new relationship element
    const ns = doc.documentElement.namespaceURI;
    const relationshipEl = ns ? doc.createElementNS(ns, element.relationship) : doc.createElement(element.relationship);
    relationshipEl.setAttribute(elementAttributes.Id, newRid);
    relationshipEl.setAttribute(elementAttributes.type, type);
    relationshipEl.setAttribute(elementAttributes.target, target);
    relationshipsElement.appendChild(relationshipEl);
    
    // Serialize and save the updated relationships file
    const serializer = new XMLSerializer();
    const newDoc = serializer.serializeToString(doc);
    zip.file(workbookRelsXmlPath, newDoc);
};

/**
 * Finds the highest relationship ID number from existing relationships in a relationships element.
 * Scans all relationship elements and extracts the numeric part from rId attributes.
 * 
 * @param relationshipsElement - The relationships XML element containing relationship elements
 * @returns The highest relationship ID number found, or 0 if none exist
 * 
 * @example
 * // If relationships contain rId1, rId3, rId7, returns 7
 * const highestRid = getHighestRelationshipId(relationshipsElement);
 */
const getHighestRelationshipId = (relationshipsElement: Element): number => {
    const relationships = relationshipsElement.getElementsByTagName(element.relationship);
    let highestRid = 0;
    
    for (let i = 0; i < relationships.length; i++) {
        const idAttr = relationships[i].getAttribute(elementAttributes.Id);
        if (idAttr && idAttr.startsWith(elementAttributes.relationshipIdPrefix)) {
            // Extract numeric part from rId (e.g., "rId5" -> 5, "rId123" -> 123)
            const ridNumber = parseInt(idAttr.substring(elementAttributes.relationshipIdPrefix.length), 10);
            if (!isNaN(ridNumber) && ridNumber > highestRid) {
                highestRid = ridNumber;
            }
        }
    }
    
    return highestRid;
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
    getSheetPathByNameFromZip,
    getReferenceFromTable,
    findTablePathFromZip,
    getCustomXmlItemNumber,
    isCustomXmlExists,
    addToContentType,
    addCustomXmlToRels,
};
