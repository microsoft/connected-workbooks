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
import { DocProps, TableData } from "../types";
import pqUtils from "./pqUtils";
import xmlInnerPartsUtils from "./xmlInnerPartsUtils";
import tableUtils from "./tableUtils";

const updateWorkbookInitialDataIfNeeded = async (
    zip: JSZip,
    docProps?: DocProps,
    tableData?: TableData,
    updateQueryTable = false
): Promise<void> => {
    await xmlInnerPartsUtils.updateDocProps(zip, docProps);
    await tableUtils.updateTableInitialDataIfNeeded(zip, tableData, updateQueryTable);
};

const updateWorkbookPowerQueryDocument = async (
    zip: JSZip,
    queryName: string,
    queryMashupDoc: string
): Promise<void> => {
    const old_base64: string | undefined = await pqUtils.getBase64(zip);

    if (!old_base64) {
        throw new Error(base64NotFoundErr);
    }

    const new_base64: string = await replaceSingleQuery(old_base64, queryName, queryMashupDoc);
    await pqUtils.setBase64(zip, new_base64);
};

const updateWorkbookSingleQueryAttributes = async (
    zip: JSZip,
    queryName: string,
    refreshOnOpen: boolean
): Promise<void> => {
    // Update connections
    const connectionsXmlString: string | undefined = await zip.file(connectionsXmlPath)?.async(textResultType);
    if (connectionsXmlString === undefined) {
        throw new Error(connectionsNotFoundErr);
    }

    const { connectionId, connectionXmlFileString } = await xmlInnerPartsUtils.updateConnections(
        connectionsXmlString,
        queryName,
        refreshOnOpen
    );
    zip.file(connectionsXmlPath, connectionXmlFileString);

    // Update sharedStrings
    const sharedStringsXmlString: string | undefined = await zip.file(sharedStringsXmlPath)?.async(textResultType);
    if (sharedStringsXmlString === undefined) {
        throw new Error(sharedStringsNotFoundErr);
    }

    const { sharedStringIndex, newSharedStrings } = await xmlInnerPartsUtils.updateSharedStrings(
        sharedStringsXmlString,
        queryName
    );
    zip.file(sharedStringsXmlPath, newSharedStrings);

    // Update sheet
    const sheetsXmlString: string | undefined = await zip.file(sheetsXmlPath)?.async(textResultType);
    if (sheetsXmlString === undefined) {
        throw new Error(sheetsNotFoundErr);
    }

    const worksheetString: string = await xmlInnerPartsUtils.updateWorksheet(
        sheetsXmlString,
        sharedStringIndex.toString()
    );
    zip.file(sheetsXmlPath, worksheetString);

    // Update tables
    await xmlInnerPartsUtils.updatePivotTablesandQueryTables(zip, queryName, refreshOnOpen, connectionId!);
};

export default {
    updateWorkbookInitialDataIfNeeded,
    updateWorkbookPowerQueryDocument,
    updateWorkbookSingleQueryAttributes,
};
