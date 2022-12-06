// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, documentUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import MashupHandler from "./mashupDocumentParser";
import { connectionsXmlPath, queryTablesPath, pivotCachesPath, docPropsCoreXmlPath, sheetsXmlPath } from "./constants";
import { DocProps, QueryInfo, docPropsAutoUpdatedElements, docPropsModifiableElements, TableData, dataTypes } from "./types";

export class WorkbookManager {
    private mashupHandler: MashupHandler = new MashupHandler();

    async generateSingleQueryWorkbook(query: QueryInfo, templateFile?: File, docProps?: DocProps, tableData?: TableData): Promise<Blob> {
        if (!query.queryMashup) {
            throw new Error("Query mashup can't be empty");
        }
        const zip =
            templateFile === undefined
                ? await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true })
                : await JSZip.loadAsync(templateFile);

        return await this.generateSingleQueryWorkbookFromZip(zip, query, docProps, tableData);
    }

    private async generateSingleQueryWorkbookFromZip(zip: JSZip, query: QueryInfo, docProps?: DocProps, tableData?: TableData): Promise<Blob> {
        await this.updatePowerQueryDocument(zip, query.queryMashup);
        await this.updateSingleQueryRefreshOnOpen(zip, query.refreshOnOpen);
        await this.updateDocProps(zip, docProps);
        if (tableData) {
            await this.addSingleQueryInitialData(zip, tableData);
        }

        return await zip.generateAsync({
            type: "blob",
            mimeType: "application/xlsx",
        });
    }

    private async updatePowerQueryDocument(zip: JSZip, queryMashup: string) {
        const old_base64 = await pqUtils.getBase64(zip);

        if (!old_base64) {
            throw new Error("Base64 string is not found in zip file");
        }

        const new_base64 = await this.mashupHandler.ReplaceSingleQuery(old_base64, queryMashup);
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

    
    private async addSingleQueryInitialData(zip: JSZip, tableData: TableData) {
        const sheetsXmlString = await zip.file(sheetsXmlPath)?.async("text");
        if (sheetsXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        const newSheet = await this.updateSheetsInitialData(sheetsXmlString, tableData);
        zip.file(sheetsXmlPath, newSheet)
    }

    private async updateSheetsInitialData(sheetsXmlString: string, tableData: TableData) {
         const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, "text/xml");
        const sheetData = sheetsDoc.getElementsByTagName("sheetData")[0];
        while (sheetData.lastChild) {
            sheetData.removeChild(sheetData.lastChild);
        }
        var rowIndex = 0;
        tableData.data.forEach((row) => {
            const newRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, "row");
            newRow.setAttribute("r", (rowIndex + 1).toString());
            newRow.setAttribute("spans", "1:" + (row.length));
            newRow.setAttribute("x14ac:dyDescent", "0.3");
            var colIndex = 0;
            row.forEach(cellContent => {
                const cell = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, "c");
                cell.setAttribute("r", String.fromCharCode(colIndex + 65) + (rowIndex + 1).toString());
                const cellData = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, "v");
                this.updateCellData(tableData, rowIndex, colIndex, cell, cellData);
                cell.appendChild(cellData);
                newRow.appendChild(cell);
                colIndex++;
            });
            sheetData.appendChild(newRow);
            rowIndex++;
        });

        sheetsDoc.getElementsByTagName("dimension")[0].setAttribute("ref", `A1:${String.fromCharCode(tableData.data[0].length + 64)}${(tableData.data.length).toString()}`);
        return serializer.serializeToString(sheetsDoc);
    }

    private updateCellData(tableData: TableData, rowIndex: number, colIndex: number, newCell: Element, cellData: Element) {
        let data = tableData.data[rowIndex][colIndex];
        if ((tableData.columnTypes[colIndex] == dataTypes.string) || (rowIndex == 0)) {
            newCell.setAttribute("t", "str");
            cellData.textContent = tableData.data[rowIndex][colIndex];
        }
        else {
            if (tableData.columnTypes[colIndex] == dataTypes.number) {          
                if (isNaN(Number(tableData.data[rowIndex][colIndex]))) {
                    data = "0";
                }
                newCell.setAttribute("t", "1");
                cellData.textContent = data;
            }

            if (tableData.columnTypes[colIndex] == dataTypes.boolean) {
                if ((tableData.data[rowIndex][colIndex] != "1") && (tableData.data[rowIndex][colIndex] != "0")) {
                    data = "0";
                }

                newCell.setAttribute("t", "b");
                cellData.textContent = data;
            }
        }
    }
    
    private async updateSingleQueryRefreshOnOpen(zip: JSZip, refreshOnOpen: boolean) {
        const connectionsXmlString = await zip.file(connectionsXmlPath)?.async("text");
        if (connectionsXmlString === undefined) {
            throw new Error("Connections were not found in template");
        }
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const refreshOnLoadValue = refreshOnOpen ? "1" : "0";
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

        queryProp.parentElement?.setAttribute("refreshOnLoad", refreshOnLoadValue);
        const connectionId = dbPr.parentElement?.getAttribute("id");
        const newConn = serializer.serializeToString(connectionsDoc);
        zip.file(connectionsXmlPath, newConn);

        if (connectionId == "-1" || !connectionId) {
            throw new Error("No connection found for Query1");
        }
        let found = false;

        // Find Query Table
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
            throw new Error("No Query Table or Pivot Table found for Query1 in given template.");
        }
    }
}
