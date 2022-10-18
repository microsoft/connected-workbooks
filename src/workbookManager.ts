// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, documentUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import MashupHandler from "./mashupDocumentParser";
import { connectionsXmlPath, workbookXmlPath, sheetsXmlPath, tableXmlPath, queryTableXmlPath, queryTablesPath, pivotCachesPath, defaults, docPropsCoreXmlPath } from "./constants";
import { DocProps, QueryInfo, docPropsAutoUpdatedElements, docPropsModifiableElements } from "./types";


export class WorkbookManager {
    private mashupHandler: MashupHandler = new MashupHandler();

    async generateSingleQueryWorkbook(query: QueryInfo, templateFile?: File, docProps?: DocProps, initialData?: string[]): Promise<Blob> {
        if (!query.queryMashup) {
            throw new Error("Query mashup can't be empty");
        }
        if (!query.queryName) {
            query.queryName = defaults.queryName;
        }
        // if (!query.initialData) {
        //     query.initialData = [['column1', 'column2'], ['111', '222']];
        // }
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
        if (query.initialData) {
            await this.addSingleQueryInitialData(zip, query.initialData);
        }
        
        return await zip.generateAsync({
            type: "blob",
            mimeType: "application/xlsx",
        });
    }

    private async addSingleQueryInitialData(zip: JSZip, initialData: string[][]) {
        //extract sheetXml
        const sheetsXmlString = await zip.file(sheetsXmlPath)?.async("text");
        if (sheetsXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, "text/xml");
        
        //edit SheetXml
        const sheetData = sheetsDoc.getElementsByTagName("sheetData")[0];
        const oldRow = sheetsDoc.getElementsByTagName("row")[0];
        const columnNames = initialData[0];
        while (sheetData.firstChild != sheetData.lastChild) {
            if (sheetData.lastChild) {
                sheetData.removeChild(sheetData.lastChild);
            }
        }
        
        documentUtils.createNewElements(sheetsDoc, columnNames, "row", "c", null, null)
        initialData.forEach((row) => {
            const newRow = oldRow.cloneNode(true);
            sheetData.appendChild(newRow);                          
        });
        if (sheetData.lastChild) {
        sheetData.removeChild(sheetData.lastChild);
        }

        const rowsArr = [...sheetsDoc.getElementsByTagName("row")];
        var rowIndex = 0;
        rowsArr.forEach((newRow) => {
            newRow.setAttribute("r", (rowIndex+1).toString());
            var colIndex = 0;
            const rowCellsArr = [...newRow.children];
            rowCellsArr.forEach((newCell) => {
                newCell.setAttribute("r", String.fromCharCode(colIndex + 65)+(rowIndex+1).toString());
                newCell.setAttribute("t", "str");
                const cellData = [...newCell.children][0];
                cellData.innerHTML = initialData[rowIndex][colIndex];
                colIndex++;
            });

            rowIndex++;
        });

        sheetsDoc.getElementsByTagName("dimension")[0].setAttribute("ref", `A1:${String.fromCharCode(initialData[0].length + 64)}${(initialData.length).toString()}`);
        const newSheet = serializer.serializeToString(sheetsDoc);
        zip.file(sheetsXmlPath, newSheet);
        
        // extract workbookXml
        const workbookXmlString = await zip.file(workbookXmlPath)?.async("text");
        if (workbookXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }

        const newParser: DOMParser = new DOMParser();
        const newSerializer = new XMLSerializer();
        const workbookDoc: Document = newParser.parseFromString(workbookXmlString, "text/xml");
        const definedName = workbookDoc.getElementsByTagName("definedName")[0];
        definedName.innerHTML = "Query1!$A$1:$" + String.fromCharCode(initialData[0].length + 64) +"$" + (initialData.length).toString();
        const newWorkbook = newSerializer.serializeToString(workbookDoc);
        zip.file(workbookXmlPath, newWorkbook);
        
        // extract TableXml
        const tableXmlString = await zip.file(tableXmlPath)?.async("text");
        if (tableXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        
        // edit tableXml columns
        const tableDoc: Document = parser.parseFromString(tableXmlString, "text/xml");
        const tablePropArr = ["id", "uniqueName", "name", "queryTableFieldId"];
        const tableElemValuesMap = new Map;
        for (var colIndex = 0; colIndex < columnNames.length; colIndex++) {
            tableElemValuesMap.set(colIndex, [(colIndex + 1).toString(), (colIndex + 1).toString(), columnNames[colIndex], (colIndex + 1).toString()]);
        }

        documentUtils.createNewElements(tableDoc, columnNames, "tableColumns", "tableColumn", tablePropArr, tableElemValuesMap);
        tableDoc.getElementsByTagName("tableColumns")[0].setAttribute("count", columnNames.length.toString());
        tableDoc.getElementsByTagName("table")[0].setAttribute("ref", `A1:${String.fromCharCode(initialData[0].length + 64)}${(initialData.length).toString()}`);
        tableDoc.getElementsByTagName("autoFilter")[0].setAttribute("ref", `A1:${String.fromCharCode(initialData[0].length + 64)}${(initialData.length).toString()}`); 
        const newTable = serializer.serializeToString(tableDoc);
        zip.file(tableXmlPath, newTable);

        //extract querytable
        const queryTableXmlString = await zip.file(queryTableXmlPath)?.async("text");
        if (queryTableXmlString === undefined) {
            throw new Error("queryTables were not found in template");
        }

        // edit querytableXml columns
        const queryTableDoc: Document = parser.parseFromString(queryTableXmlString, "text/xml");
        const queryTablePropArr = ["id", "name", "tableColumnId"];
        const queryTableElemValuesMap = new Map;
        for (var fieldIndex = 0; fieldIndex < columnNames.length; fieldIndex++) {
            queryTableElemValuesMap.set(fieldIndex, [(fieldIndex + 1).toString(), columnNames[fieldIndex], (fieldIndex + 1).toString()]);
        }

        documentUtils.createNewElements(queryTableDoc, columnNames, "queryTableFields", "queryTableField", queryTablePropArr, queryTableElemValuesMap);
        queryTableDoc.getElementsByTagName("queryTableFields")[0].setAttribute("count", columnNames.length.toString()); 
        queryTableDoc.getElementsByTagName("queryTableRefresh")[0].setAttribute("nextId", (columnNames.length + 1).toString())
        const newQueryTable = serializer.serializeToString(queryTableDoc);
        zip.file(queryTableXmlPath, newQueryTable);
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

        // Update RefreshOnOpen
        //queryProp.parentElement?.setAttribute("refreshOnLoad", refreshOnLoadValue);
        
        // Update query details to match queryName
        dbPr.parentElement?.setAttribute("name", `Query - ${queryName}`);
        dbPr.parentElement?.setAttribute("description", `Connection to the ${queryName} query in the workbook.`);
        dbPr.setAttribute("connection", `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=${queryName};`);
        dbPr.setAttribute("command", `SELECT * FROM [${queryName}]`);
        
        const connectionId = dbPr.parentElement?.getAttribute("id");
        const newConn = serializer.serializeToString(connectionsDoc);
        zip.file(connectionsXmlPath, newConn);

        if (connectionId == "-1" || !connectionId) {
            throw new Error(`No connection found for ${queryName}`);
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
                //element.setAttribute("refreshOnLoad", refreshOnLoadValue);
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
                //element.parentElement!.setAttribute("refreshOnLoad", refreshOnLoadValue);
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
