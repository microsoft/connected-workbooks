// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import {
    base64NotFoundErr,
    connectionsXmlPath,
    textResultType,
    connectionsNotFoundErr,
    sharedStringsXmlPath,
    sharedStringsNotFoundErr,
    sheetsXmlPath,
    sheetsNotFoundErr,
    xmlTextResultType,
} from "./constants";
import { addConnectionOnlyQueries, replaceSingleQuery } from "./mashupDocumentParser";
import { FileConfigs, TableData } from "../types";
import pqUtils from "./pqUtils";
import xmlInnerPartsUtils from "./xmlInnerPartsUtils";
import tableUtils from "./tableUtils";

const updateWorkbookDataAndConfigurations = async (zip: JSZip, fileConfigs?: FileConfigs, tableData?: TableData, updateQueryTable = false): Promise<void> => {
    await xmlInnerPartsUtils.updateDocProps(zip, fileConfigs?.docProps);
    if (fileConfigs?.templateFile === undefined) {
        // If we are using our base template, we need to clear label info
        await xmlInnerPartsUtils.clearLabelInfo(zip);
    }
    await tableUtils.updateTableInitialDataIfNeeded(zip, tableData, updateQueryTable);
};

const updateWorkbookPowerQueryDocument = async (zip: JSZip, loadedQueryName: string, queryMashupDoc: string, connectionOnlyQueryNames?: string[]): Promise<void> => {
    const old_base64: string | undefined = await pqUtils.getBase64(zip);
    if (!old_base64) {
        throw new Error(base64NotFoundErr);
    }
    // The mashupDoc contains a default query, we replace that query with the loaded query 
    let updated_base64: string = await replaceSingleQuery(old_base64, loadedQueryName, queryMashupDoc);
    
    // If connection-only queries were given, add them to the mashupDoc
    updated_base64 = await addConnectionOnlyQueriesIfNeeded(updated_base64, connectionOnlyQueryNames);          

    await pqUtils.setBase64(zip, updated_base64);
};

const addConnectionOnlyQueriesIfNeeded = async(base64: string, connectionOnlyQueryNames?:string[]):Promise<string> => {
    if (!connectionOnlyQueryNames || (connectionOnlyQueryNames.length == 0))
    {
        return base64;
    } 

    return await addConnectionOnlyQueries(base64, connectionOnlyQueryNames);
};



const updateWorkbookSingleQueryAttributes = async (zip: JSZip, queryName: string, refreshOnOpen: boolean): Promise<void> => {
    // Update connections
    const connectionsXmlString: string | undefined = await zip.file(connectionsXmlPath)?.async(textResultType);
    if (connectionsXmlString === undefined) {
        throw new Error(connectionsNotFoundErr);
    }

    const { connectionId, connectionXmlFileString } = xmlInnerPartsUtils.updateConnections(connectionsXmlString, queryName, refreshOnOpen);
    zip.file(connectionsXmlPath, connectionXmlFileString);

    // Update sharedStrings
    const sharedStringsXmlString: string | undefined = await zip.file(sharedStringsXmlPath)?.async(textResultType);
    if (sharedStringsXmlString === undefined) {
        throw new Error(sharedStringsNotFoundErr);
    }

    const { sharedStringIndex, newSharedStrings } = xmlInnerPartsUtils.updateSharedStrings(sharedStringsXmlString, queryName);
    zip.file(sharedStringsXmlPath, newSharedStrings);

    // Update sheet
    const sheetsXmlString: string | undefined = await zip.file(sheetsXmlPath)?.async(textResultType);
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const worksheetString: string = xmlInnerPartsUtils.updateWorksheet(sheetsXmlString, sharedStringIndex.toString());
    zip.file(sheetsXmlPath, worksheetString);

    // Update tables
    await xmlInnerPartsUtils.updatePivotTablesandQueryTables(zip, queryName, refreshOnOpen, connectionId!);
};

const addConnectionOnlyQueriesToWorkbook = async (zip: JSZip, connectionOnlyQueryNames: string[]): Promise<void> => {
    // Update connections
    let connectionsXmlString: string | undefined = await zip.file(connectionsXmlPath)?.async(textResultType);
    if (connectionsXmlString === undefined) {
        throw new Error(connectionsNotFoundErr);
    }

    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    let connectionsDoc: Document = parser.parseFromString(connectionsXmlString, xmlTextResultType);
    connectionOnlyQueryNames.forEach(async (queryName: string) => { 
        connectionsDoc = await xmlInnerPartsUtils.addNewConnection(connectionsDoc, queryName);
    });
    
    connectionsXmlString = serializer.serializeToString(connectionsDoc);
    zip.file(connectionsXmlPath, connectionsXmlString);
    
};

export default {
    updateWorkbookDataAndConfigurations,
    updateWorkbookPowerQueryDocument,
    updateWorkbookSingleQueryAttributes,
    addConnectionOnlyQueriesToWorkbook,
};
