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
    tableXmlPath,
    defaults,
    tablesFolderPath,
} from "./constants";
import { replaceSingleQuery } from "./mashupDocumentParser";
import { FileConfigs, TableData, TemplateSettings } from "../types";
import pqUtils from "./pqUtils";
import tableUtils from "./tableUtils";
import xmlInnerPartsUtils from "./xmlInnerPartsUtils";
import documentUtils from "./documentUtils";

const updateWorkbookDataAndConfigurations = async (zip: JSZip, fileConfigs?: FileConfigs, tableData?: TableData, updateQueryTable = false): Promise<void> => {
    let sheetName: string = defaults.sheetName;
    let tablePath: string = tableXmlPath;
    let sheetPath: string = sheetsXmlPath;

    if (fileConfigs?.templateFile !== undefined) {
        const templateSettings: TemplateSettings | undefined = fileConfigs?.templateSettings;

        // Getting the sheet id based on location in the workbook
        if (templateSettings?.sheetName !== undefined) {
            const sheetLocation = await xmlInnerPartsUtils.getSheetPathByNameFromZip(zip, templateSettings.sheetName);
            sheetName = templateSettings.sheetName;
            sheetPath = "xl/" + sheetLocation;
        }

        // Getting the table location based on which table has the same name as the one in the fileConfigs
        // If no table name is provided, we will use the default one
        if (templateSettings?.tableName !== undefined) {
            tablePath = tablesFolderPath + await xmlInnerPartsUtils.findTablePathFromZip(zip, templateSettings?.tableName);
        }
    }
    
    // Getting the table start and end location string from the table path
    // If no table path is provided, we will consider A1 as the start location
    let cellRangeRef: string = "A1";
    if (fileConfigs?.templateFile != null) {
        cellRangeRef = await xmlInnerPartsUtils.getReferenceFromTable(zip, tablePath)
    }

   if (tableData) {
        cellRangeRef += `:${documentUtils.getCellReferenceRelative(tableData.columnNames.length - 1, tableData.rows.length + 1)}`;
    }
    
    await xmlInnerPartsUtils.updateDocProps(zip, fileConfigs?.docProps);
    if (fileConfigs?.templateFile === undefined) {
        // If we are using our base template, we need to clear label info
        await xmlInnerPartsUtils.clearLabelInfo(zip);
    }
    await tableUtils.updateTableInitialDataIfNeeded(zip, cellRangeRef, sheetPath, tablePath, sheetPath, tableData, updateQueryTable);
};

const updateWorkbookPowerQueryDocument = async (zip: JSZip, queryName: string, queryMashupDoc: string): Promise<void> => {
    const old_base64: string | null = await pqUtils.getBase64(zip);

    if (!old_base64) {
        throw new Error(base64NotFoundErr);
    }

    const new_base64: string = await replaceSingleQuery(old_base64, queryName, queryMashupDoc);
    await pqUtils.setBase64(zip, new_base64);
};

const updateWorkbookSingleQueryAttributes = async (zip: JSZip, queryName: string, refreshOnOpen: boolean, sheetName?: string): Promise<void> => {
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
    let sheetPath: string = sheetsXmlPath;
    if (sheetName !== undefined) {
        const sheetLocation = await xmlInnerPartsUtils.getSheetPathByNameFromZip(zip, sheetName);
            sheetPath = "xl/" + sheetLocation;
    }
    
    const sheetsXmlString: string | undefined = await zip.file(sheetPath)?.async(textResultType);
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const worksheetString: string = xmlInnerPartsUtils.updateWorksheet(sheetsXmlString, sharedStringIndex.toString());
    zip.file(sheetPath, worksheetString);

    // Update tables
    await xmlInnerPartsUtils.updatePivotTablesandQueryTables(zip, queryName, refreshOnOpen, connectionId!);
};

export default {
    updateWorkbookDataAndConfigurations,
    updateWorkbookPowerQueryDocument,
    updateWorkbookSingleQueryAttributes,
};
