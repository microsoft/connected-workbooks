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
} from "./constants";
import { replaceSingleQuery } from "./mashupDocumentParser";
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

const updateWorkbookPowerQueryDocument = async (zip: JSZip, queryName: string, queryMashupDoc: string): Promise<void> => {
    const old_base64: string | undefined = await pqUtils.getBase64(zip);

    if (!old_base64) {
        throw new Error(base64NotFoundErr);
    }

    const new_base64: string = await replaceSingleQuery(old_base64, queryName, queryMashupDoc);
    await pqUtils.setBase64(zip, new_base64);
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

    connectionOnlyQueryNames.forEach(async (queryName: string) => { 
        connectionsXmlString = await xmlInnerPartsUtils.addNewConnection(connectionsXmlString!, queryName);
    });
    
};

export default {
    updateWorkbookDataAndConfigurations,
    updateWorkbookPowerQueryDocument,
    updateWorkbookSingleQueryAttributes,
    addConnectionOnlyQueriesToWorkbook,
};
