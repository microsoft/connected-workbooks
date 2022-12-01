// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, documentUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import MashupHandler from "./mashupDocumentParser";
import { connectionsXmlPath, queryTablesPath, pivotCachesPath, docPropsCoreXmlPath, defaults, sharedStringsXmlPath, sheetsXmlPath } from "./constants";
import { DocProps, QueryInfo, docPropsAutoUpdatedElements, docPropsModifiableElements } from "./types";

export class WorkbookManager {
    private mashupHandler: MashupHandler = new MashupHandler();

    async generateSingleQueryWorkbook(query: QueryInfo, templateFile?: File, docProps?: DocProps): Promise<Blob> {
        if (!query.queryMashup) {
            throw new Error("Query mashup can't be empty");
        }
        if (!query.queryName) {
            query.queryName = defaults.queryName;
        }
        const zip =
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
            type: "blob",
            mimeType: "application/xlsx",
        });
    }

    private async updatePowerQueryDocument(zip: JSZip, queryName: string, queryMashup: string) {
        const old_base64 = await pqUtils.getBase64(zip);

        if (!old_base64) {
            throw new Error("Base64 string is not found in zip file");
        }

        const new_base64 = await this.mashupHandler.ReplaceSingleQuery(old_base64, queryName, queryMashup);
        await pqUtils.setBase64(zip, new_base64);
    }

    private async updateDocProps(zip: JSZip, docProps: DocProps = {}) {
        const { doc, properties } = await documentUtils.getDocPropsProperties(zip);

        //set auto updated elements
        const docPropsAutoUpdatedElementsArr = Object.keys(docPropsAutoUpdatedElements) as Array<
            keyof typeof docPropsAutoUpdatedElements
        >;

        const nowTime = new Date().toISOString();

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

        const serializer = new XMLSerializer();
        const newDoc = serializer.serializeToString(doc);
        zip.file(docPropsCoreXmlPath, newDoc);
    }

    private async updateSingleQueryAttributes(zip: JSZip, queryName: string, refreshOnOpen: boolean) {
        //Update connections
        const connectionsXmlString = await zip.file(connectionsXmlPath)?.async("text");
        if (connectionsXmlString === undefined) {
            throw new Error("Connections were not found in template");
        }  
        
        const {connectionId, connectionString} = await this.updateConnections(connectionsXmlString, queryName, refreshOnOpen);
        zip.file(connectionsXmlPath, connectionString);
        
        //Update sharedStrings
        const sharedStringsXmlString = await zip.file(sharedStringsXmlPath)?.async("text");
        if (sharedStringsXmlString === undefined) {
            throw new Error("SharedStrings were not found in template");
        }
        const {sharedStringIndex, newSharedStrings} = await this.updateSharedStrings(sharedStringsXmlString, queryName);
        zip.file(sharedStringsXmlPath, newSharedStrings);
        
        //Update sheet
        const sheetsXmlString = await zip.file(sheetsXmlPath)?.async("text");
        if (sheetsXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        const worksheetString = await this.updateWorksheet(sheetsXmlString, sharedStringIndex.toString());
        zip.file(sheetsXmlPath, worksheetString);
        
        //Update tables
        await this.updatePivotTablesandQueryTables(zip, queryName, refreshOnOpen, connectionId!);  
    }

    private async updateConnections(connectionsXmlString: string, queryName: string, refreshOnOpen: boolean) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const refreshOnLoadValue = refreshOnOpen ? "1" : "0";
        const connectionsDoc: Document = parser.parseFromString(connectionsXmlString, "text/xml");
        const connectionsProperties = connectionsDoc.getElementsByTagName("dbPr");
        const dbPr = connectionsProperties[0];
        dbPr.setAttribute("refreshOnLoad", refreshOnLoadValue);
        
        // Update query details to match queryName
        dbPr.parentElement?.setAttribute("name", `Query - ${queryName}`);
        dbPr.parentElement?.setAttribute("description", `Connection to the '${queryName}' query in the workbook.`);
        dbPr.setAttribute("connection", `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=${queryName};`);
        dbPr.setAttribute("command",`SELECT * FROM [${queryName}]`);
        const connectionId = dbPr.parentElement?.getAttribute("id");
        const connectionString = serializer.serializeToString(connectionsDoc);

        if (connectionId === null) {
            throw new Error(`No connection found for ${queryName}`);
        }

        return {connectionId, connectionString};
    }

    private async updateSharedStrings(sharedStringsXmlString: string, queryName: string) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sharedStringsDoc: Document = parser.parseFromString(sharedStringsXmlString, "text/xml");
        const tItems = sharedStringsDoc.getElementsByTagName("t");
        let t = null;
        let sharedStringIndex = tItems.length;
        if (tItems && tItems.length) {
            for (let i = 0; i < tItems.length; i++) {
                if (tItems[i].innerHTML === queryName) {
                    t = tItems[i];
                    sharedStringIndex = i + 1;
                } 
            }
        }
        const sst = sharedStringsDoc.getElementsByTagName("sst")[0];
        if (t === null) {
            if (!sst) {
                throw new Error("No shared string was found!");
            }   
            if (sharedStringsDoc.documentElement.namespaceURI) {
                const tElement = sharedStringsDoc.createElementNS(sharedStringsDoc.documentElement.namespaceURI, "t");
                tElement.textContent = queryName;
                const siElement = sharedStringsDoc.createElementNS(sharedStringsDoc.documentElement.namespaceURI, "si");
                siElement.appendChild(tElement);
                sharedStringsDoc.getElementsByTagName("sst")[0].appendChild(siElement);
                        }
            const value = sst.getAttribute("count");
            if (value) {
                sst.setAttribute("count", (parseInt(value)+1).toString()); 
            }
            const uniqueValue = sst.getAttribute("uniqueCount");
            if (uniqueValue) {
                sst.setAttribute("uniqueCount", (parseInt(uniqueValue)+1).toString()); 
            }
        }
        const newSharedStrings = serializer.serializeToString(sharedStringsDoc);
        return {sharedStringIndex, newSharedStrings};

}

    private async updateWorksheet(sheetsXmlString: string, sharedStringIndex: string) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, "text/xml");
        sheetsDoc.getElementsByTagName("v")[0].innerHTML = sharedStringIndex.toString();
        const newSheet = serializer.serializeToString(sheetsDoc);
        return newSheet;
    }

    private async updatePivotTablesandQueryTables(zip: JSZip, queryName: string, refreshOnOpen: boolean, connectionId: string) {
        // Find Query Table
        let found = false;
        const queryTablePromises: Promise<{
            path: string;
            queryTableXmlString: string;
        }>[] = [];
        zip.folder(queryTablesPath)?.forEach(async (relativePath, queryTableFile) => {
            queryTablePromises.push(
                (() => {
                    return queryTableFile.async("text").then((queryTableString) => {
                        return {
                            path: relativePath,
                            queryTableXmlString: queryTableString,
                        };
                    });
                })()
            );
        });
        
        (await Promise.all(queryTablePromises)).forEach(({ path, queryTableXmlString }) => {
            const {isTableUpdated, newTable} = this.updateTable(queryTableXmlString, connectionId, refreshOnOpen, "queryTable");
            zip.file(queryTablesPath + path, newTable);
            if (isTableUpdated) {
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
            if (relativePath.startsWith("pivotCacheDefinition")) {
                pivotCachePromises.push(
                    (() => {
                        return pivotCacheFile.async("text").then((pivotCacheString) => {
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
            const {isTableUpdated, newTable} = this.updateTable(pivotCacheXmlString, connectionId, refreshOnOpen, "cacheSource");
            zip.file(pivotCachesPath + path, newTable);
            if (isTableUpdated) {
                found = true;
            }
        });
        if (!found) {
            throw new Error(`No Query Table or Pivot Table found for ${queryName} in given template.`);
        }
    }

    private updateTable(tableXmlString: string, connectionId: string, refreshOnOpen: boolean, elementName: string) {
        const refreshOnLoadValue = refreshOnOpen ? "1" : "0";
        let isTableUpdated = false;
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const pivotCacheDoc: Document = parser.parseFromString(tableXmlString, "text/xml");
        let element = pivotCacheDoc.getElementsByTagName(elementName)[0];
        var newTable = "";
        if (element.getAttribute("connectionId") == connectionId) {
            if (elementName === "cacheSource") {
                element = element.parentElement!;
            }
            element.setAttribute("refreshOnLoad", refreshOnLoadValue);
            newTable = serializer.serializeToString(pivotCacheDoc);
            isTableUpdated = true;
        }
        return {isTableUpdated, newTable};
    }

} 
