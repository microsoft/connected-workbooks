// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, xmlPartsUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import {
    defaults,
    emptyQueryMashupErr,
    blobFileType,
    application,
    templateWithInitialDataErr,
    tableNotFoundErr,
} from "./utils/constants";
import { DocProps, QueryInfo, TableData, Grid, TableDataParser, FileConfigs } from "./types";
import TableDataParserFactory from "./TableDataParserFactory";
import { generateSingleQueryMashup } from "./generators";
import { extractTableValues } from "./utils/htmlUtils";

const generateSingleQueryWorkbook = async (
    query: QueryInfo,
    initialDataGrid?: Grid,
    fileConfigs?: FileConfigs
): Promise<Blob> => {
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
        templateFile === undefined
            ? await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true })
            : await JSZip.loadAsync(templateFile);

    const tableData: TableData | undefined = await parseInitialDataGrid(initialDataGrid);

    return await generateSingleQueryWorkbookFromZip(zip, query, fileConfigs?.docProps, tableData);
};

const generateTableWorkbookFromHtml = async (htmlTable: HTMLTableElement, docProps?: DocProps): Promise<Blob> => {
    const gridData = extractTableValues(htmlTable);
    return await generateTableWorkbookFromGrid({ gridData: gridData, promoteHeaders: false }, docProps);
};

const generateTableWorkbookFromGrid = async (initialDataGrid: Grid, docProps?: DocProps): Promise<Blob> => {
    const zip: JSZip = await JSZip.loadAsync(WorkbookTemplate.SIMPLE_BLANK_TABLE_TEMPLATE, { base64: true });
    const tableData: TableData | undefined = await parseInitialDataGrid(initialDataGrid);
    if (tableData === undefined) {
        throw new Error(tableNotFoundErr);
    }

    await xmlPartsUtils.updateWorkbookInitialDataIfNeeded(zip, docProps, tableData);

    return await zip.generateAsync({
        type: blobFileType,
        mimeType: application,
    });
};

const parseInitialDataGrid = async (initialDataGrid?: Grid): Promise<TableData | undefined> => {
    if (!initialDataGrid) {
        return undefined;
    }

    const parser: TableDataParser = TableDataParserFactory.createParser(initialDataGrid);
    const tableData: TableData | undefined = parser.parseToTableData(initialDataGrid);

    return tableData;
};

const generateSingleQueryWorkbookFromZip = async (
    zip: JSZip,
    query: QueryInfo,
    docProps?: DocProps,
    tableData?: TableData
): Promise<Blob> => {
    if (!query.queryName) {
        query.queryName = defaults.queryName;
    }

    await xmlPartsUtils.updateWorkbookPowerQueryDocument(
        zip,
        query.queryName,
        generateSingleQueryMashup(query.queryName, query.queryMashup)
    );
    await xmlPartsUtils.updateWorkbookSingleQueryAttributes(zip, query.queryName, query.refreshOnOpen);
    await xmlPartsUtils.updateWorkbookInitialDataIfNeeded(zip, docProps, tableData, true /*updateQueryTable*/);

    return await zip.generateAsync({
        type: blobFileType,
        mimeType: application,
    });
};

const downloadWorkbook = (file: Blob, filename: string): void => {
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

export default {
    generateSingleQueryWorkbook,
    generateTableWorkbookFromHtml,
    generateTableWorkbookFromGrid,
    downloadWorkbook,
};
