// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, documentUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import MashupHandler from "./mashupDocumentParser";
import { connectionsXmlPath, queryTablesPath, pivotCachesPath, docPropsCoreXmlPath, defaults, sharedStringsXmlPath, sheetsXmlPath, emptyQueryMashupErr, blobFileType, application, base64NotFoundErr, textResultType, connectionsNotFoundErr, sharedStringsNotFoundErr, sheetsNotFoundErr, trueValue, falseValue, xmlTextResultType, element, elementAttributes, elementAttributesValues, pivotCachesPathPrefix, emptyValue, queryAndPivotTableNotFoundErr } from "./constants";
import { DocProps, QueryInfo, docPropsAutoUpdatedElements, docPropsModifiableElements } from "./types";
import arrayUtils, { ArrayReader } from "./utils/arrayUtils";

export class WorkbookManager {
    private mashupHandler: MashupHandler = new MashupHandler();

    async generateSingleQueryWorkbook(query: QueryInfo, templateFile?: File, docProps?: DocProps): Promise<Blob> {
        if (!query.queryMashup) {
            throw new Error(emptyQueryMashupErr);
        }

        if (!query.queryName) {
            query.queryName = defaults.queryName;
        }

        const zip: JSZip =
            templateFile === undefined
                ? await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true })
                : await JSZip.loadAsync(templateFile);

        return await this.generateSingleQueryWorkbookFromZip(zip, query, docProps);
    }

    private async generateSingleQueryWorkbookFromZip(zip: JSZip, query: QueryInfo, docProps?: DocProps): Promise<Blob> {
        if (!query.queryName) {
            query.queryName = defaults.queryName;
        }

        await this.updatePowerQueryDocument(zip, query.queryName, query.queryMashup);
        await this.updateSingleQueryAttributes(zip, query.queryName, query.refreshOnOpen);
        await this.updateDocProps(zip, docProps);

        return await zip.generateAsync({
            type: blobFileType,
            mimeType: application,
        });
    }

    private async updatePowerQueryDocument(zip: JSZip, queryName: string, queryMashup: string) {
        const old_base64: string | undefined = await pqUtils.getBase64(zip);

        if (!old_base64) {
            throw new Error(base64NotFoundErr);
        }

        const new_base64: string = await this.mashupHandler.ReplaceSingleQuery(old_base64, queryName, queryMashup);
        await pqUtils.setBase64(zip, new_base64);
    }

    private async updateDocProps(zip: JSZip, docProps: DocProps = {}) {
        const { doc, properties } = await documentUtils.getDocPropsProperties(zip);

        //set auto updated elements
        const docPropsAutoUpdatedElementsArr: ("created" | "modified")[] = Object.keys(docPropsAutoUpdatedElements) as Array<
            keyof typeof docPropsAutoUpdatedElements
        >;

        const nowTime: string = new Date().toISOString();

        docPropsAutoUpdatedElementsArr.forEach((tag) => {
            documentUtils.createOrUpdateProperty(doc, properties, docPropsAutoUpdatedElements[tag], nowTime);
        });

        //set modifiable elements
        const docPropsModifiableElementsArr = Object.keys(docPropsModifiableElements) as Array<
            keyof typeof docPropsModifiableElements
        >;

        docPropsModifiableElementsArr
            .map((key) => ({
                name: docPropsModifiableElements[key],
                value: docProps[key],
            }))
            .forEach((kvp) => {
                documentUtils.createOrUpdateProperty(doc, properties, kvp.name!, kvp.value);
            });

        const serializer: XMLSerializer = new XMLSerializer();
        const newDoc: string = serializer.serializeToString(doc);
        zip.file(docPropsCoreXmlPath, newDoc);
    }

    private async updateSingleQueryAttributes(zip: JSZip, queryName: string, refreshOnOpen: boolean) {
        //Update connections
        const connectionsXmlString: string | undefined = await zip.file(connectionsXmlPath)?.async(textResultType);
        if (connectionsXmlString === undefined) {
            throw new Error(connectionsNotFoundErr);
        }

        const { connectionId, connectionXmlFileString } = await this.updateConnections(connectionsXmlString, queryName, refreshOnOpen);
        zip.file(connectionsXmlPath, connectionXmlFileString);

        //Update sharedStrings
        const sharedStringsXmlString: string | undefined = await zip.file(sharedStringsXmlPath)?.async(textResultType);
        if (sharedStringsXmlString === undefined) {
            throw new Error(sharedStringsNotFoundErr);
        }

        const { sharedStringIndex, newSharedStrings } = await this.updateSharedStrings(sharedStringsXmlString, queryName);
        zip.file(sharedStringsXmlPath, newSharedStrings);

        //Update sheet
        const sheetsXmlString: string | undefined = await zip.file(sheetsXmlPath)?.async(textResultType);
        if (sheetsXmlString === undefined) {
            throw new Error(sheetsNotFoundErr);
        }

        const worksheetString: string = await this.updateWorksheet(sheetsXmlString, sharedStringIndex.toString());
        zip.file(sheetsXmlPath, worksheetString);

        //Update tables
        await this.updatePivotTablesandQueryTables(zip, queryName, refreshOnOpen, connectionId!);
    }

    private async updateConnections(connectionsXmlString: string, queryName: string, refreshOnOpen: boolean) {
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
    }

    private async updateSharedStrings(sharedStringsXmlString: string, queryName: string) {
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
    }

    private async updateWorksheet(sheetsXmlString: string, sharedStringIndex: string) {
        const parser: DOMParser = new DOMParser();
        const serializer: XMLSerializer = new XMLSerializer();
        const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, xmlTextResultType);
        sheetsDoc.getElementsByTagName(element.cellValue)[0].innerHTML = sharedStringIndex.toString();
        const newSheet: string = serializer.serializeToString(sheetsDoc);

        return newSheet;
    }

    private async updatePivotTablesandQueryTables(zip: JSZip, queryName: string, refreshOnOpen: boolean, connectionId: string) {
        // Find Query Table
        let found: boolean = false;
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
            const { isQueryTableUpdated, newQueryTable } = this.updateQueryTable(queryTableXmlString, connectionId, refreshOnOpen);
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
            const { isPivotTableUpdated, newPivotTable } = this.updatePivotTable(pivotCacheXmlString, connectionId, refreshOnOpen);
            zip.file(pivotCachesPath + path, newPivotTable);
            if (isPivotTableUpdated) {
                found = true;
            }
        });
        if (!found) {
            throw new Error(queryAndPivotTableNotFoundErr);
        }
    }

    private updateQueryTable(tableXmlString: string, connectionId: string, refreshOnOpen: boolean) {
        const refreshOnLoadValue: string = refreshOnOpen ? trueValue : falseValue;
        let isQueryTableUpdated: boolean = false;
        const parser: DOMParser = new DOMParser();
        const serializer: XMLSerializer = new XMLSerializer();
        const queryTableDoc: Document = parser.parseFromString(tableXmlString, xmlTextResultType);
        const queryTable: Element = queryTableDoc.getElementsByTagName(element.queryTable)[0];
        var newQueryTable: string = emptyValue;
        if (queryTable.getAttribute(elementAttributes.connectionId) == connectionId) {
            queryTable.setAttribute(elementAttributes.refreshOnLoad, refreshOnLoadValue);
            newQueryTable = serializer.serializeToString(queryTableDoc);
            isQueryTableUpdated = true;
        }

        return { isQueryTableUpdated, newQueryTable };
    }

    private updatePivotTable(tableXmlString: string, connectionId: string, refreshOnOpen: boolean) {
        const refreshOnLoadValue: string = refreshOnOpen ? trueValue : falseValue;
        let isPivotTableUpdated: boolean = false;
        const parser: DOMParser = new DOMParser();
        const serializer: XMLSerializer = new XMLSerializer();
        const pivotCacheDoc: Document = parser.parseFromString(tableXmlString, xmlTextResultType);
        let cacheSource: Element = pivotCacheDoc.getElementsByTagName(element.cacheSource)[0];
        var newPivotTable: string = emptyValue;
        if (cacheSource.getAttribute(elementAttributes.connectionId) == connectionId) {
            cacheSource = cacheSource.parentElement!;
            cacheSource.setAttribute(elementAttributes.refreshOnLoad, refreshOnLoadValue);
            newPivotTable = serializer.serializeToString(pivotCacheDoc);
            isPivotTableUpdated = true;
        }

        return { isPivotTableUpdated, newPivotTable };
    }

    public async getMQueryData(zipFilePath: string) {
        var fs = require("fs");

        const mashupHandler = new MashupHandler();
        const data = fs.readFileSync(zipFilePath);
        const zipFile = await JSZip.loadAsync(data);
        const originalBase64Str = await pqUtils.getBase64(zipFile);

        const { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } =
            mashupHandler.getPackageComponents(originalBase64Str!);
        const packageZip: JSZip = await JSZip.loadAsync(packageOPC);
        const section1m = await mashupHandler.getSection1m(packageZip);

        console.log(version, packageOPC, permissionsSize, permissions, metadata, endBuffer);

        //extract metadataXml
        const mashupArray: ArrayReader = new arrayUtils.ArrayReader(metadata.buffer);
        const metadataVersion: Uint8Array = mashupArray.getBytes(4);
        const metadataXmlSize: number = mashupArray.getInt32();
        const metadataXml: Uint8Array = mashupArray.getBytes(metadataXmlSize);

        //parse metdataXml
        const textDecoder: TextDecoder = new TextDecoder();
        const metadataString: string = textDecoder.decode(metadataXml);
        const parser: DOMParser = new DOMParser();
        const serializer: XMLSerializer = new XMLSerializer();
        const parsedMetadata: Document = parser.parseFromString(metadataString, xmlTextResultType);
        const entries = parsedMetadata.getElementsByTagName(element.entry);
        if (entries && entries.length) {
            for (let i = 0; i < entries.length; i++) {
                const entry: Element = entries[i];
                const entryAttributes: NamedNodeMap = entry.attributes;
                const entryAttributesArr: Attr[] = [...entryAttributes];
                const entryProp: Attr | undefined = entryAttributesArr.find((prop) => {
                    return prop?.name === elementAttributes.type;
                });

                if (entryProp?.nodeValue == elementAttributes.fillTarget) {
                    console.log(entryProp?.nodeValue);
                    console.log(entryProp?.value);

                    const entryValues: Attr | undefined = entryAttributesArr.find((prop) => {
                        return prop?.name === elementAttributes.value;
                    });
                    console.log(entryValues);
                }
            }
        }

        // extract query name from query with regex
        const queryName = section1m.match(/shared\s(.+?) =/)?.[1];
        console.log(queryName);

        // extract connection ID from connections using the query name

        // extract metadata from query table using the connection ID
    }

    public async getQueryInfo(zipFilePath: string): Promise<string> {
        var fs = require("fs");
        const data = fs.readFileSync(zipFilePath);
        const zipFile = await JSZip.loadAsync(data);
        const originalBase64Str = await pqUtils.getBase64(zipFile);

        const mashupHandler = new MashupHandler();
        const { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } =
            mashupHandler.getPackageComponents(originalBase64Str!);
        // extract section1m
        const packageZip: JSZip = await JSZip.loadAsync(packageOPC);
        const section1m = await mashupHandler.getSection1m(packageZip);
        return section1m;
    }

} 
