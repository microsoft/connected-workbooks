// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, documentUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import MashupHandler from "./mashupDocumentParser";
import {
    connectionsXmlPath,
    queryTablesPath,
    pivotCachesPath,
    docPropsCoreXmlPath, defaults, sharedStringsXmlPath, sheetsXmlPath, emptyQueryMashupErr, blobFileType, application, base64NotFoundErr, textResultType, connectionsNotFoundErr, sharedStringsNotFoundErr, sheetsNotFoundErr, trueValue, falseValue, xmlTextResultType, element, elementAttributes, elementAttributesValues, pivotCachesPathPrefix, emptyValue, queryAndPivotTableNotFoundErr,
    queryTableXmlPath,
    tableXmlPath,
    workbookXmlPath,
} from "./constants";
import {
    DocProps,
    QueryInfo,
    docPropsAutoUpdatedElements,
    docPropsModifiableElements,
    TableData,
    dataTypes,
    Grid,
    ColumnMetadata,
} from "./types";
import TableDataParserFactory from "./TableDataParserFactory";

export class WorkbookManager {
    private mashupHandler: MashupHandler = new MashupHandler();

    async generateSingleQueryWorkbook(
        query: QueryInfo,
        initialDataGrid?: Grid,
        templateFile?: File,
        docProps?: DocProps
    ): Promise<Blob> {
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
        if (templateFile !== undefined && initialDataGrid !== undefined) {
            throw new Error("Cannot receive template file with initial data");
        }
        const tableData = await this.parseInitialDataGrid(initialDataGrid);
        return await this.generateSingleQueryWorkbookFromZip(zip, query, docProps, tableData);
    }

    private async parseInitialDataGrid(initialDataGrid?: Grid): Promise<TableData | undefined> {
        if (!initialDataGrid) {
            return undefined;
        }
        const parser = TableDataParserFactory.createParser(initialDataGrid);
        const tableData = parser.parseToTableData(initialDataGrid);
        return tableData;
    }


    private async generateSingleQueryWorkbookFromZip(
        zip: JSZip,
        query: QueryInfo,
        docProps?: DocProps,
        tableData?: TableData
    ): Promise<Blob> {
        if (!query.queryName) {
            query.queryName = defaults.queryName;
        }

        await this.updatePowerQueryDocument(zip, query.queryName, query.queryMashup);
        await this.updateSingleQueryAttributes(zip, query.queryName, query.refreshOnOpen);
        await this.updateDocProps(zip, docProps);
        if (tableData) {
            await this.addSingleQueryInitialData(zip, tableData);
        }
        return await zip.generateAsync({
            type: blobFileType,
            mimeType: application,
        });
    }

    private async updatePowerQueryDocument(zip: JSZip, queryName: string, queryMashup: string) {
        const old_base64: string|undefined = await pqUtils.getBase64(zip);

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

    private async addSingleQueryInitialData(zip: JSZip, tableData: TableData) {
        const sheetsXmlString = await zip.file(sheetsXmlPath)?.async("text");
        if (sheetsXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        const newSheet = await this.updateSheetsInitialData(sheetsXmlString, tableData);
        zip.file(sheetsXmlPath, newSheet);

        const queryTableXmlString = await zip.file(queryTableXmlPath)?.async("text");
        if (queryTableXmlString === undefined) {
            throw new Error("Query Table was not found in template");
        }
        const newQueryTable = await this.updateQueryTablesInitialData(queryTableXmlString, tableData);
        zip.file(queryTableXmlPath, newQueryTable);

        const tableXmlString = await zip.file(tableXmlPath)?.async("text");
        if (tableXmlString === undefined) {
            throw new Error("Table was not found in template");
        }
        const newTable = await this.updateTablesInitialData(tableXmlString, tableData);
        zip.file(tableXmlPath, newTable);

        const workbookXmlString = await zip.file(workbookXmlPath)?.async("text");
        if (workbookXmlString === undefined) {
            throw new Error("Sheets were not found in template");
        }
        const newWorkbook = await this.updateWorkbookInitialData(workbookXmlString, tableData);
        zip.file(workbookXmlPath, newWorkbook);
    }

    private async updateTablesInitialData(tableXmlString: string, tableData: TableData) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const tableDoc: Document = parser.parseFromString(tableXmlString, "text/xml");
        const tableColumns = tableDoc.getElementsByTagName("tableColumns")[0];
        tableColumns.textContent = "";
        tableData.columnMetadata.forEach((col: ColumnMetadata, columnIndex: number) => {
            const tableColumn = tableDoc.createElementNS(tableDoc.documentElement.namespaceURI, "tableColumn");
            tableColumn.setAttribute("id", (columnIndex + 1).toString());
            tableColumn.setAttribute("uniqueName", (columnIndex + 1).toString());
            tableColumn.setAttribute("name", col.name);
            tableColumn.setAttribute("queryTableFieldId", (columnIndex + 1).toString());
            tableColumns.appendChild(tableColumn);
        });

        tableColumns.setAttribute("count", tableData.columnMetadata.length.toString());
        tableDoc
            .getElementsByTagName("table")[0]
            .setAttribute(
                "ref",
                `A1:${documentUtils.getCellReference(
                    tableData.columnMetadata.length - 1,
                    tableData.data.length + 1
                )}`.replace("$", "")
            );
        tableDoc
            .getElementsByTagName("autoFilter")[0]
            .setAttribute(
                "ref",
                `A1:${documentUtils.getCellReference(
                    tableData.columnMetadata.length - 1,
                    tableData.data.length + 1
                )}`.replace("$", "")
            );
        return serializer.serializeToString(tableDoc);
    }

    private async updateWorkbookInitialData(workbookXmlString: string, tableData: TableData, queryName?: string) {
        const newParser: DOMParser = new DOMParser();
        const newSerializer = new XMLSerializer();
        const workbookDoc: Document = newParser.parseFromString(workbookXmlString, "text/xml");
        var definedName = workbookDoc.getElementsByTagName("definedName")[0];
        const prefix = queryName === undefined ? "Query1" : queryName;
        definedName.textContent =
            prefix +
            `!$A$1:$${documentUtils.getCellReference(tableData.columnMetadata.length - 1, tableData.data.length + 1)}`;
        return newSerializer.serializeToString(workbookDoc);
    }

    private async updateQueryTablesInitialData(queryTableXmlString: string, tableData: TableData) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const queryTableDoc: Document = parser.parseFromString(queryTableXmlString, "text/xml");
        const queryTableFields = queryTableDoc.getElementsByTagName("queryTableFields")[0];
        queryTableFields.textContent = "";
        tableData.columnMetadata.forEach((col: ColumnMetadata, columnIndex: number) => {
            const queryTableField = queryTableDoc.createElementNS(
                queryTableDoc.documentElement.namespaceURI,
                "queryTableField"
            );
            queryTableField.setAttribute("id", (columnIndex + 1).toString());
            queryTableField.setAttribute("name", col.name);
            queryTableField.setAttribute("tableColumnId", (columnIndex + 1).toString());
            queryTableFields.appendChild(queryTableField);
        });
        queryTableFields.setAttribute("count", tableData.columnMetadata.length.toString());
        queryTableDoc
            .getElementsByTagName("queryTableRefresh")[0]
            .setAttribute("nextId", (tableData.columnMetadata.length + 1).toString());
        return serializer.serializeToString(queryTableDoc);
    }

    private async updateSheetsInitialData(sheetsXmlString: string, tableData: TableData) {
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const sheetsDoc: Document = parser.parseFromString(sheetsXmlString, "text/xml");
        const sheetData = sheetsDoc.getElementsByTagName("sheetData")[0];
        sheetData.textContent = "";
        var rowIndex = 0;
        const columnRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, "row");
        columnRow.setAttribute("r", (rowIndex + 1).toString());
        columnRow.setAttribute("spans", "1:" + tableData.columnMetadata.length);
        columnRow.setAttribute("x14ac:dyDescent", "0.3");
        tableData.columnMetadata.forEach((col, colIndex) => {
            columnRow.appendChild(documentUtils.createCell(sheetsDoc, colIndex, rowIndex, dataTypes.string, col.name));
        });
        sheetData.appendChild(columnRow);
        rowIndex++;
        tableData.data.forEach((row) => {
            const newRow = sheetsDoc.createElementNS(sheetsDoc.documentElement.namespaceURI, "row");
            newRow.setAttribute("r", (rowIndex + 1).toString());
            newRow.setAttribute("spans", "1:" + row.length);
            newRow.setAttribute("x14ac:dyDescent", "0.3");
            row.forEach((cellContent, colIndex) => {
                newRow.appendChild(
                    documentUtils.createCell(
                        sheetsDoc,
                        colIndex,
                        rowIndex,
                        tableData.columnMetadata[colIndex].type,
                        cellContent
                    )
                );
            });
            sheetData.appendChild(newRow);
            rowIndex++;
        });

        sheetsDoc
            .getElementsByTagName("dimension")[0]
            .setAttribute("ref", documentUtils.getTableReference(tableData.data[0].length - 1, tableData.data.length));
        return serializer.serializeToString(sheetsDoc);
    }

    private async updateSingleQueryAttributes(zip: JSZip, queryName: string, refreshOnOpen: boolean) {
        //Update connections
        const connectionsXmlString: string|undefined = await zip.file(connectionsXmlPath)?.async(textResultType);
        if (connectionsXmlString === undefined) {
            throw new Error(connectionsNotFoundErr);
        }  
        
        const {connectionId, connectionXmlFileString } = await this.updateConnections(connectionsXmlString, queryName, refreshOnOpen);
        zip.file(connectionsXmlPath, connectionXmlFileString );
        
        //Update sharedStrings
        const sharedStringsXmlString: string|undefined = await zip.file(sharedStringsXmlPath)?.async(textResultType);
        if (sharedStringsXmlString === undefined) {
            throw new Error(sharedStringsNotFoundErr);
        }
        
        const {sharedStringIndex, newSharedStrings} = await this.updateSharedStrings(sharedStringsXmlString, queryName);
        zip.file(sharedStringsXmlPath, newSharedStrings);
        
        //Update sheet
        const sheetsXmlString: string|undefined = await zip.file(sheetsXmlPath)?.async(textResultType);
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
        dbPr.setAttribute(elementAttributes.command,elementAttributesValues.connectionCommand(queryName));
        const connectionId: string | null | undefined = dbPr.parentElement?.getAttribute(elementAttributes.id);
        const connectionXmlFileString: string  = serializer.serializeToString(connectionsDoc);

        if (connectionId === null) {
            throw new Error(connectionsNotFoundErr);
        }

        return {connectionId, connectionXmlFileString};
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
        let textElement: Element|null = null;
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

            const value: string|null = sharedStringsTable.getAttribute(elementAttributes.count);
            if (value) {
                sharedStringsTable.setAttribute(elementAttributes.count, (parseInt(value)+1).toString()); 
            }

            const uniqueValue: string|null = sharedStringsTable.getAttribute(elementAttributes.uniqueCount);
            if (uniqueValue) {
                sharedStringsTable.setAttribute(elementAttributes.uniqueCount, (parseInt(uniqueValue)+1).toString()); 
            }
        }
        const newSharedStrings: string = serializer.serializeToString(sharedStringsDoc);
        
        return {sharedStringIndex, newSharedStrings};
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
            const {isQueryTableUpdated, newQueryTable} = this.updateQueryTable(queryTableXmlString, connectionId, refreshOnOpen);
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
            const {isPivotTableUpdated, newPivotTable} = this.updatePivotTable(pivotCacheXmlString, connectionId, refreshOnOpen);
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

        return {isQueryTableUpdated, newQueryTable};
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

        return {isPivotTableUpdated, newPivotTable};
    }

} 
