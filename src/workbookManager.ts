// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, xmlPartsUtils } from "./utils";
import { SIMPLE_BLANK_TABLE_TEMPLATE, SIMPLE_QUERY_WORKBOOK_TEMPLATE } from "./workbookTemplate";
import { defaults, emptyQueryMashupErr, blobFileType, application, templateWithInitialDataErr, tableNotFoundErr } from "./utils/constants";
import { DocProps, QueryInfo, TableData, Grid, FileConfigs } from "./types";
import { generateSingleQueryMashup } from "./generators";
import { extractTableValues } from "./utils/htmlUtils";
import { parseToTableData } from "./gridParser";

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

    const tableData: TableData | undefined = await parseInitialDataGrid(initialDataGrid);

    return await generateSingleQueryWorkbookFromZip(zip, query, fileConfigs?.docProps, tableData);
};

export const generateTableWorkbookFromHtml = async (htmlTable: HTMLTableElement, docProps?: DocProps): Promise<Blob> => {
    const gridData = extractTableValues(htmlTable);
    return await generateTableWorkbookFromGrid({ data: gridData, config: { promoteHeaders: true } }, docProps);
};

export const generateTableWorkbookFromGrid = async (grid: Grid, docProps?: DocProps): Promise<Blob> => {
    const zip: JSZip = await JSZip.loadAsync(SIMPLE_BLANK_TABLE_TEMPLATE, { base64: true });
    const tableData: TableData | undefined = await parseInitialDataGrid(grid);
    if (tableData === undefined) {
        throw new Error(tableNotFoundErr);
    }

    await xmlPartsUtils.updateWorkbookInitialDataIfNeeded(zip, docProps, tableData);

    return await zip.generateAsync({
        type: blobFileType,
        mimeType: application,
    });
};

const parseInitialDataGrid = async (grid?: Grid): Promise<TableData | undefined> => {
    if (!grid) {
        return undefined;
    }

    const tableData: TableData | undefined = parseToTableData(grid);

    return tableData;
};

const generateSingleQueryWorkbookFromZip = async (zip: JSZip, query: QueryInfo, docProps?: DocProps, tableData?: TableData): Promise<Blob> => {
    if (!query.queryName) {
        query.queryName = defaults.queryName;
    }

    await xmlPartsUtils.updateWorkbookPowerQueryDocument(zip, query.queryName, generateSingleQueryMashup(query.queryName, query.queryMashup));
    await xmlPartsUtils.updateWorkbookSingleQueryAttributes(zip, query.queryName, query.refreshOnOpen);
    await xmlPartsUtils.updateWorkbookInitialDataIfNeeded(zip, docProps, tableData, true /*updateQueryTable*/);

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
