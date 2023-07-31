// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, xmlPartsUtils, htmlUtils, gridUtils } from "./utils";
import { SIMPLE_BLANK_TABLE_TEMPLATE, SIMPLE_QUERY_WORKBOOK_TEMPLATE } from "./workbookTemplate";
import {
    defaults,
    emptyQueryMashupErr,
    blobFileType,
    application,
    templateWithInitialDataErr,
    tableNotFoundErr,
    templateFileNotSupportedErr,
} from "./utils/constants";
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
    if (fileConfigs?.templateFile !== undefined) {
        throw new Error(templateFileNotSupportedErr);
    }
    const gridData = htmlUtils.extractTableValues(htmlTable);
    return await generateTableWorkbookFromGrid({ data: gridData, config: { promoteHeaders: true } }, fileConfigs);
};

export const generateTableWorkbookFromGrid = async (grid: Grid, fileConfigs?: FileConfigs): Promise<Blob> => {
    if (fileConfigs?.templateFile !== undefined) {
        throw new Error(templateFileNotSupportedErr);
    }
    const zip: JSZip = await JSZip.loadAsync(SIMPLE_BLANK_TABLE_TEMPLATE, { base64: true });
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
