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
        const refreshOnLoadValue = refreshOnOpen ? "1" : "0";
        const connectionId = await this.editConnections(zip, queryName, refreshOnLoadValue);
        const sharedStringId = (await this.editSharedStrings(zip, queryName)).toString();
        await this.editWorksheet(zip, sharedStringId);
        await this.editPivotTable(zip, queryName, refreshOnLoadValue, connectionId);  
    }

    private async editConnections(zip: JSZip, queryName: string, refreshOnLoadValue: string) {
        const connectionsXmlString = await zip.file(connectionsXmlPath)?.async("text");
        if (connectionsXmlString === undefined) {
            throw new Error("Connections were not found in template");
        }
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const connectionsDoc: Document = parser.parseFromString(connectionsXmlString, "text/xml");

        const connectionsProperties = connectionsDoc.getElementsByTagName("dbPr");

        const dbPr = connectionsProperties[0];
        const connectionsAttributes = dbPr.attributes;
        const connectionsAttributesArr = [...connectionsAttributes];

        const queryProp = connectionsAttributesArr.find((prop) => {
            return prop?.name === "command" && prop.nodeValue === "SELECT * FROM [Query1]";
        });

        if (!queryProp) {
            throw new Error("No query was found!");
        }
        
        // Update RefreshOnOpen
        queryProp.parentElement?.setAttribute("refreshOnLoad", refreshOnLoadValue);
        
        // Update query details to match queryName
        dbPr.parentElement?.setAttribute("name", `Query - ${queryName}`);
        dbPr.parentElement?.setAttribute("description", `Connection to the '${queryName}' query in the workbook.`);
        dbPr.setAttribute("connection", `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=${queryName};`);
        dbPr.setAttribute("command",`SELECT * FROM [${queryName}]`);
        const connectionId = dbPr.parentElement?.getAttribute("id");
        const newConn = serializer.serializeToString(connectionsDoc);
        zip.file(connectionsXmlPath, newConn);

        if (connectionId == "-1" || !connectionId) {
            throw new Error(`No connection found for ${queryName}`);
        }
        return connectionId;
    }

    private async editSharedStrings(zip: JSZip, queryName: string) {
        // edit shared string
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sharedStringsXmlString = await zip.file(sharedStringsXmlPath)?.async("text");
        if (sharedStringsXmlString === undefined) {
            throw new Error("SharedStrings were not found in template");
        }
        const sharedStringsDoc: Document = parser.parseFromString(sharedStringsXmlString, "text/xml");
        const tItems = sharedStringsDoc.getElementsByTagName("t");
        let t = null;
        let sharedStringIndex = tItems.length;
        if (tItems && tItems.length) {
            for (let i = 0; i < tItems.length; i++) {
                if (tItems[i].innerHTML === queryName) {
                    t = tItems[i];
                    sharedStringIndex = i;
                } 
            }
        }
        if (t === null) {
            const sst = sharedStringsDoc.getElementsByTagName("sst")[0];
            if (!sst) {
                throw new Error("No shared string was found!");
            }           
            const oldSi = sst.firstChild;
            if (oldSi) {
                sst.appendChild(oldSi.cloneNode(true));
                sharedStringsDoc.getElementsByTagName("t")[tItems.length - 1].innerHTML = queryName;
            }     
        }
        const newSharedStrings = serializer.serializeToString(sharedStringsDoc);
        zip.file(sharedStringsXmlPath, newSharedStrings);
        return sharedStringIndex;
    }

    private async editWorksheet(zip: JSZip, sharedStringIndex: string) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sheetsXmlString = await zip.file(sheetsXmlPath)?.async("text");
        if (sheetsXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, "text/xml");
        //edit columnName to correct shared string element 
        sheetsDoc.getElementsByTagName("v")[0].innerHTML = sharedStringIndex.toString();
        const newSheet = serializer.serializeToString(sheetsDoc);
        zip.file(sheetsXmlPath, newSheet);
    }

    private async editPivotTable(zip: JSZip, queryName: string, refreshOnLoadValue: string, connectionId: string) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        // Find Query Table
        const queryTablePromises: Promise<{
            path: string;
            queryTableXmlString: string;
        }>[] = [];
        let found = false;
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
            const queryTableDoc: Document = parser.parseFromString(queryTableXmlString, "text/xml");
            const element = queryTableDoc.getElementsByTagName("queryTable")[0];
            if (element.getAttribute("connectionId") == connectionId) {
                element.setAttribute("refreshOnLoad", refreshOnLoadValue);
                const newQT = serializer.serializeToString(queryTableDoc);
                zip.file(queryTablesPath + path, newQT);
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
            const pivotCacheDoc: Document = parser.parseFromString(pivotCacheXmlString, "text/xml");
            const element = pivotCacheDoc.getElementsByTagName("cacheSource")[0];
            if (element.getAttribute("connectionId") == connectionId) {
                element.parentElement!.setAttribute("refreshOnLoad", refreshOnLoadValue);
                const newPC = serializer.serializeToString(pivotCacheDoc);
                zip.file(pivotCachesPath + path, newPC);
                found = true;
            }
        });
        if (!found) {
            throw new Error(`No Query Table or Pivot Table found for ${queryName} in given template.`);
        }
    }
} 
