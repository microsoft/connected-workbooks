// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, xmlPartsUtils, htmlUtils, gridUtils } from "./utils";
import { SIMPLE_BLANK_TABLE_TEMPLATE, SIMPLE_QUERY_WORKBOOK_TEMPLATE } from "./workbookTemplate";
import { defaults, emptyQueryMashupErr, blobFileType, application, templateWithInitialDataErr, tableNotFoundErr, headers, OFU } from "./utils/constants";
import { QueryInfo, TableData, Grid, FileConfigs } from "./types";
import { generateSingleQueryMashup } from "./generators";

export const generateSingleQueryWorkbook = async (query: QueryInfo, initialDataGrid?: Grid, fileConfigs?: FileConfigs): Promise<Blob> => {
    if (!query.queryMashup) {
        throw new Error(emptyQueryMashupErr);
    }

    if (!query.queryName) {
        query.queryName = defaults.queryName;
    }

    const templateFile: File | undefined = fileConfigs?.templateFile;
    if (templateFile !== undefined && initialDataGrid !== undefined) {
        throw new Error(templateWithInitialDataErr);
    }

    pqUtils.validateQueryName(query.queryName);

    const zip: JSZip =
        templateFile === undefined ? await JSZip.loadAsync(SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true }) : await JSZip.loadAsync(templateFile);

    const tableData = initialDataGrid ? gridUtils.parseToTableData(initialDataGrid) : undefined;

    return await generateSingleQueryWorkbookFromZip(zip, query, fileConfigs, tableData);
};

export const generateTableWorkbookFromHtml = async (htmlTable: HTMLTableElement, fileConfigs?: FileConfigs): Promise<Blob> => {
    const gridData = htmlUtils.extractTableValues(htmlTable);
    return await generateTableWorkbookFromGrid({ data: gridData, config: { promoteHeaders: true } }, fileConfigs);
};

export const generateTableWorkbookFromGrid = async (grid: Grid, fileConfigs?: FileConfigs): Promise<Blob> => {
    const zip: JSZip =
        fileConfigs?.templateFile === undefined
            ? await JSZip.loadAsync(SIMPLE_BLANK_TABLE_TEMPLATE, { base64: true })
            : await JSZip.loadAsync(fileConfigs.templateFile);

    const tableData = gridUtils.parseToTableData(grid);
    if (tableData === undefined) {
        throw new Error(tableNotFoundErr);
    }

    await xmlPartsUtils.updateWorkbookDataAndConfigurations(zip, fileConfigs, tableData);

    return await zip.generateAsync({
        type: blobFileType,
        mimeType: application,
    });
};

const generateSingleQueryWorkbookFromZip = async (zip: JSZip, query: QueryInfo, fileConfigs?: FileConfigs, tableData?: TableData): Promise<Blob> => {
    if (!query.queryName) {
        query.queryName = defaults.queryName;
    }

    await xmlPartsUtils.updateWorkbookPowerQueryDocument(zip, query.queryName, generateSingleQueryMashup(query.queryName, query.queryMashup));
    await xmlPartsUtils.updateWorkbookSingleQueryAttributes(zip, query.queryName, query.refreshOnOpen);
    await xmlPartsUtils.updateWorkbookDataAndConfigurations(zip, fileConfigs, tableData, true /*updateQueryTable*/);

    return await zip.generateAsync({
        type: blobFileType,
        mimeType: application,
    });
};

export const downloadWorkbook = (file: Blob, filename: string): void => {
    const nav = window.navigator as any;
    if (nav.msSaveOrOpenBlob)
        // IE10+
        nav.msSaveOrOpenBlob(file, filename);
    else {
        // Others
        const a = document.createElement("a");
        const url = URL.createObjectURL(file);
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 0);
    }
};

export const openInExcelWeb = async (file: Blob, filename?: string, allowTyping?: boolean): Promise<void> => {
    try {
        const url = await getExcelForWebUrl(file, filename, allowTyping);
        window.open(url, "_blank");
    } catch (error) {
        console.error("An error occurred:", error);
    }
};

export const getExcelForWebUrl = async (file: Blob, filename?: string, allowTyping?: boolean): Promise<string> => {
    // Check if the file exists
    if (file.size < 0) {
        throw new Error("File is empty");
    }

    // Read the content of the Excel file into a buffer
    const fileContent = file;
    const fileNameGuid = new Date().getTime().toString() + (filename ? "_" + filename : "") + ".xlsx";

    // Parse allowTyping parameter
    const allowTypingParam = allowTyping ? 1 : 0;

    try {
        // Send the POST request to the desired endpoint using Fetch
        const response = await fetch(`${OFU.PostUrl}${fileNameGuid}&${OFU.WdOrigin}=${OFU.OpenInExcelOririgin}`, {
            method: "POST",
            headers: headers,
            body: fileContent,
        });

        // Check if the response is successful
        if (response.ok) {
            // if upload was successful - open the file in a new tab
            return `${OFU.ViewUrl}${fileNameGuid}&${OFU.AllowTyping}=${allowTypingParam}&${OFU.WdOrigin}=${OFU.OpenInExcelOririgin}`;
        } else {
            throw new Error(`File upload failed. Status code: ${response.status}`);
        }
    } catch (error) {
        throw new Error(`An error occurred: ${error}`);
    }
};
